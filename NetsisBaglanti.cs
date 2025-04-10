using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks; // SemaphoreSlim.WaitAsync için eklendi (opsiyonel)
using Microsoft.Extensions.Logging; // Loglama için eklendi (projenize göre ayarlayın)
using NetOpenX50;

namespace NetAi
{
    public class NetsisConnectionPool : IDisposable
    {
        private readonly ConcurrentBag<NetsisConnection> _connections;
        private readonly SemaphoreSlim _poolSemaphore;
        private int _totalConnections; // Toplam oluşturulan bağlantı sayısı (Interlocked ile yönetilecek)
        private int _activeConnections; // Aktif kullanımdaki bağlantılar (Interlocked ile yönetilecek)
        private readonly object _creationLock = new object(); // Sadece yeni bağlantı oluşturma anında çakışmayı önlemek için
        private bool _disposed;

        private readonly string _server;
        private readonly string _database;
        private readonly string _dbUser;
        private readonly string _dbPassword;
        private readonly string _netsisUser;
        private readonly string _netsisPassword;
        private readonly int _branchCode;
        private readonly int _maxPoolSize;
        private readonly ILogger<NetsisConnectionPool> _logger; // Loglama için

        // Interlocked ile sayaçları okumak daha güvenli olabilir
        public string PoolStatus =>
            $"Aktif: {Volatile.Read(ref _activeConnections)}, Havuzda: {_connections.Count}, Toplam: {Volatile.Read(ref _totalConnections)}, Max: {_maxPoolSize}";

        public NetsisConnectionPool(string server, string database, string dbUser, string dbPassword,
                                  string netsisUser, string netsisPassword, int branchCode,
                                  int maxPoolSize = 1, ILogger<NetsisConnectionPool> logger = null) // Logger eklendi
        {
            if (maxPoolSize <= 0) throw new ArgumentOutOfRangeException(nameof(maxPoolSize), "MaxPoolSize pozitif bir değer olmalıdır.");

            _server = server;
            _database = database;
            _dbUser = dbUser;
            _dbPassword = dbPassword;
            _netsisUser = netsisUser;
            _netsisPassword = netsisPassword;
            _branchCode = branchCode;
            _maxPoolSize = maxPoolSize;
            // Null logger yerine dummy logger kullanmak null check'leri azaltır
            _logger = logger ?? Microsoft.Extensions.Logging.Abstractions.NullLogger<NetsisConnectionPool>.Instance;
            _connections = new ConcurrentBag<NetsisConnection>();
            _poolSemaphore = new SemaphoreSlim(maxPoolSize, maxPoolSize); // Başlangıç ve maksimum değer
            _disposed = false;

            // Başlangıçta havuzu doldurmak yerine isteğe bağlı bırakılabilir veya
            // az sayıda bağlantı eklenebilir. Şimdilik boş başlatıyoruz.
            _logger.LogInformation("NetsisConnectionPool başlatıldı. MaxPoolSize: {MaxPoolSize}", _maxPoolSize);
        }

        public NetsisConnection GetConnection(TimeSpan? timeout = null) // Timeout eklendi
        {
            CheckDisposed();

            // Varsayılan timeout (örneğin 30 saniye)
            var waitTimeout = timeout ?? TimeSpan.FromSeconds(30);

            _logger.LogTrace("Bağlantı bekleniyor... {PoolStatus}", PoolStatus);

            // Havuzdan bir slot (yer) almaya çalış
            if (!_poolSemaphore.Wait(waitTimeout))
            {
                throw new TimeoutException($"Netsis bağlantı havuzundan '{waitTimeout}' süresi içinde bağlantı alınamadı. {PoolStatus}");
            }

            _logger.LogTrace("Semaphore alındı. Bağlantı alınıyor/oluşturuluyor... {PoolStatus}", PoolStatus);
            NetsisConnection connection = null;
            bool success = false;
            try
            {
                while (true) // Geçersiz bağlantı alınırsa tekrar denemek için döngü
                {
                    if (_connections.TryTake(out connection))
                    {
                        _logger.LogTrace("Havuzdan mevcut bağlantı alındı.");
                        // Bağlantıyı doğrula ve aç
                        if (connection.ValidateAndOpen(_logger))
                        {
                            Interlocked.Increment(ref _activeConnections);
                            _logger.LogDebug("GetConnection (Havuzdan): {PoolStatus}", PoolStatus);
                            success = true;
                            return connection; // Başarılı, bağlantıyı döndür
                        }
                        else
                        {
                            // Bağlantı geçersiz, dispose et ve tekrar dene
                            _logger.LogWarning("Havuzdan alınan bağlantı geçersiz, dispose ediliyor. {PoolStatus}", PoolStatus);
                            connection.Dispose();
                            Interlocked.Decrement(ref _totalConnections);
                            // Semaphore'u hemen serbest bırakıp tekrar beklememek için döngüye devam et
                            continue;
                        }
                    }
                    else // Havuz boş, yeni bağlantı oluşturmayı dene
                    {
                        // Zaten semaphore aldığımız için _totalConnections < _maxPoolSize kontrolüne
                        // teorik olarak gerek yok ama double check için zararı olmaz.
                        // Ancak aynı anda birden fazla thread'in buraya girip total'i kontrol etmesi
                        // race condition yaratabilir. Semaphore bunu zaten engelliyor.

                        _logger.LogTrace("Havuz boş, yeni bağlantı oluşturulacak.");
                        // Sadece bir thread'in aynı anda bağlantı oluşturmasını sağlamak için ek kilit
                        // Bu, COM başlatma çakışmalarını önleyebilir.
                        lock (_creationLock)
                        {
                            // Tekrar kontrol et, başka bir thread oluşturmuş olabilir
                            if (_connections.TryTake(out connection))
                            {
                                if (connection.ValidateAndOpen(_logger))
                                {
                                    Interlocked.Increment(ref _activeConnections);
                                    _logger.LogDebug("GetConnection (Yeni oluşturulmuştu, başkası aldı): {PoolStatus}", PoolStatus);
                                    success = true;
                                    return connection;
                                }
                                else
                                {
                                    _logger.LogWarning("Yeni oluşturulan bağlantı (başka thread tarafından) geçersiz, dispose ediliyor.");
                                    connection.Dispose();
                                    Interlocked.Decrement(ref _totalConnections);
                                    continue; // Döngüye devam et, tekrar dene
                                }
                            }

                            // Yeni bağlantı oluşturma mantığı
                            if (Volatile.Read(ref _totalConnections) < _maxPoolSize)
                            {
                                connection = CreateNewConnection();
                                Interlocked.Increment(ref _totalConnections);
                                // Yeni oluşturulan bağlantı zaten açık gelmiyor, ValidateAndOpen açacak
                                if (connection.ValidateAndOpen(_logger))
                                {
                                    Interlocked.Increment(ref _activeConnections);
                                    _logger.LogInformation("Yeni bağlantı oluşturuldu ve alındı: {PoolStatus}", PoolStatus);
                                    success = true;
                                    return connection; // Başarılı
                                }
                                else
                                {
                                    _logger.LogError("Yeni oluşturulan bağlantı açılamadı/geçersiz, dispose ediliyor.");
                                    connection.Dispose(); // Hata durumunda hemen dispose et
                                    Interlocked.Decrement(ref _totalConnections);
                                    // Başarısız oldu, döngü başa dönecek ve muhtemelen başka bir bağlantı deneyecek
                                    // veya timeout'a düşecek.
                                    throw new InvalidOperationException($"Yeni oluşturulan Netsis bağlantısı doğrulanamadı. {PoolStatus}");
                                }
                            }
                            else
                            {
                                // Bu durumun semaphore nedeniyle olmaması gerekir, ama olursa diye...
                                _logger.LogError("Kritik Hata: Semaphore alındı ama _totalConnections >= _maxPoolSize! {PoolStatus}", PoolStatus);
                                throw new InvalidOperationException("Havuz dolu görünüyor ancak semaphore alınabildi. Mantıksal hata.");
                            }
                        } // lock _creationLock
                    }
                } // while(true)
            }
            finally
            {
                // Eğer bağlantı alınamadıysa (exception vb.) veya geçersiz olduğu için
                // dispose edildiyse, semaphore'u serbest bırak. Başarılı durumda semaphore
                // ReleaseConnection'da serbest bırakılacak.
                if (!success)
                {
                    _logger.LogWarning("GetConnection başarısız oldu veya geçersiz bağlantı nedeniyle semaphore serbest bırakılıyor.");
                    _poolSemaphore.Release();
                }
            }
        }


        public void ReleaseConnection(NetsisConnection connection)
        {
            if (connection == null) return;

            // Havuz dispose edildiyse veya bağlantı zaten dispose edildiyse işlem yapma
            if (_disposed || connection.IsDisposed)
            {
                // Eğer bağlantı dispose edildiyse ve havuzdan alındıysa sayaçları ve semaphore'u ayarla
                if (!connection.IsDisposed) // Bu kontrol gereksiz gibi ama garanti olsun
                {
                    connection.Dispose(); // Emin olmak için tekrar çağır
                }
                // Bu bağlantı artık havuzun bir parçası değilse sayaçları güncelleme.
                // Ancak GetConnection'dan geldiyse güncellemek gerekir.
                // Bu kısmı yönetmek karmaşıklaşabilir. Şimdilik sadece loglayalım.
                _logger.LogTrace("ReleaseConnection: Bağlantı veya havuz zaten dispose edilmiş.");
                // NOT: Eğer GetConnection başarılı olduysa ama Release'den ÖNCE
                // bağlantı/havuz dispose edildiyse semaphore eksik kalabilir.
                // Bu durumu engellemek için GetConnection'daki try-finally önemli.
                // Dispose sırasında semaphore release edilmemeli.
                return;
            }


            Interlocked.Decrement(ref _activeConnections); // Aktif sayısını azalt

            try
            {
                // Bağlantıyı havuza koymadan önce kapat (isteğe bağlı, açık da bırakılabilir)
                // Eğer açık bırakılacaksa ValidateAndOpen'in tekrar açma mantığı gözden geçirilmeli.
                connection.NetRS?.Kapat(); // Kapatma hatasını yoksayabiliriz belki?
                _logger.LogTrace("Bağlantı kapatıldı (NetRS.Kapat).");
            }
            catch (Exception ex)
            {
                // Kapatma hatası kritik olmayabilir, loglamak yeterli.
                _logger.LogWarning(ex, "NetRS.Kapat sırasında hata oluştu, ancak bağlantı havuza ekleniyor.");
            }

            _connections.Add(connection);
            _logger.LogDebug("ReleaseConnection: Bağlantı havuza eklendi. {PoolStatus}", PoolStatus);


            try
            {
                _poolSemaphore.Release(); // Havuza bir bağlantı döndüğü için semaphore'u serbest bırak
                _logger.LogTrace("Semaphore serbest bırakıldı.");
            }
            catch (SemaphoreFullException ex)
            {
                // Bu durum, havuz mantığında bir hata olduğunu gösterir (Release sayısı Wait sayısını geçti).
                _logger.LogCritical(ex, "Kritik Hata: SemaphoreFullException! Havuza kapasitesinden fazla bağlantı bırakılmaya çalışıldı. {PoolStatus}", PoolStatus);
                // Hatalı durumu düzeltmek için bağlantıyı dispose edip total'i azaltabiliriz.
                try
                {
                    connection.Dispose();
                    Interlocked.Decrement(ref _totalConnections);
                    _logger.LogWarning("SemaphoreFullException sonrası fazla bağlantı dispose edildi.");
                }
                catch (Exception disposeEx)
                {
                    _logger.LogError(disposeEx, "SemaphoreFullException sonrası bağlantı dispose edilirken hata.");
                }

            }
        }

        public T ExecuteWithConnection<T>(Func<NetRS, T> action, TimeSpan? timeout = null)
        {
            var connection = GetConnection(timeout);
            try
            {
                // Execute işleminin kendisi de hata verebilir
                return action(connection.NetRS);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ExecuteWithConnection (Func) sırasında hata oluştu.");
                // Hata durumunda bağlantı hala geçerli mi? Bu duruma özel işlem gerekebilir.
                // Örneğin, bağlantıyı dispose edip havuzdan eksiltmek.
                // Şimdilik sadece hatayı tekrar fırlatıyoruz.
                throw;
            }
            finally
            {
                ReleaseConnection(connection);
            }
        }

        public void ExecuteWithConnection(Action<NetRS> action, TimeSpan? timeout = null)
        {
            var connection = GetConnection(timeout);
            try
            {
                action(connection.NetRS);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ExecuteWithConnection (Action) sırasında hata oluştu.");
                throw;
            }
            finally
            {
                ReleaseConnection(connection);
            }
        }


        private NetsisConnection CreateNewConnection()
        {
            CheckDisposed();
            _logger.LogTrace("Yeni Netsis bağlantısı oluşturuluyor...");

            Kernel kernel = null;
            Sirket company = null;
            NetRS netRS = null;

            try
            {
                // Kernel oluşturma pahalı olabilir, dikkatli kullanılmalı.
                kernel = new Kernel();
                // _server parametresi yeniSirket metodunda yok, VT Tipi ile belirtiliyor.
                // Eğer _server gerekiyorsa, Netsis API'sine göre kullanımını kontrol edin.
                company = kernel.yeniSirket(
                    TVTTipi.vtMSSQL, // Varsayılan, gerekirse değiştirin
                    _database,
                    _dbUser,
                    _dbPassword,
                    _netsisUser,
                    _netsisPassword,
                    _branchCode
                );

                if (company == null) // Bağlantı başarısız olabilir
                {
                    throw new InvalidOperationException($"Netsis şirket bağlantısı kurulamadı (kernel.yeniSirket null döndü). DB: {_database}, User: {_netsisUser}");
                }

                netRS = kernel.yeniNetRS(company);
                if (netRS == null) // NetRS oluşturma başarısız olabilir
                {
                    throw new InvalidOperationException($"Netsis NetRS nesnesi oluşturulamadı (kernel.yeniNetRS null döndü).");
                }

                // Yeni bağlantıyı kapalı durumda oluşturuyoruz, ValidateAndOpen açacak.
                // netRS.Kapat(); // Bu satıra gerek yok, ValidateAndOpen yapacak.

                _logger.LogInformation("Yeni Netsis bağlantı nesneleri (Kernel, Sirket, NetRS) başarıyla oluşturuldu.");
                return new NetsisConnection(kernel, company, netRS, _logger); // Logger'ı ilet
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Netsis bağlantısı oluşturulurken hata oluştu.");
                // Hata durumunda oluşturulmuş olabilecek COM nesnelerini temizle
                if (netRS != null) { try { Marshal.ReleaseComObject(netRS); } catch { } }
                if (company != null) { try { Marshal.ReleaseComObject(company); } catch { } }
                if (kernel != null)
                {
                    try { kernel.FreeNetsisLibrary(); } catch { }
                    try { Marshal.ReleaseComObject(kernel); } catch { }
                }
                throw; // Hatayı tekrar fırlat
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this); // Finalizer'ı çağırmaya gerek kalmadı
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                _logger.LogInformation("NetsisConnectionPool dispose ediliyor... {PoolStatus}", PoolStatus);
                _disposed = true; // Önce işaretle ki yeni bağlantı alınamasın

                // Semaphore'u dispose et (bekleyenleri serbest bırakmaz, hata fırlatır)
                _poolSemaphore.Dispose();

                // Havuzdaki tüm bağlantıları alıp dispose et
                while (_connections.TryTake(out var conn))
                {
                    try
                    {
                        conn.Dispose();
                        // Dispose sırasında _totalConnections'ı azaltmak race condition'a neden olabilir
                        // ve zaten havuz yok oluyor. Sayaçları sıfırlamak yeterli.
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "Havuzdaki bir bağlantı dispose edilirken hata oluştu.");
                    }
                }
                _logger.LogInformation("NetsisConnectionPool dispose edildi.");
            }
            // Yönetilmeyen kaynakları burada temizle (varsa)

            _totalConnections = 0;
            _activeConnections = 0;
            _disposed = true; // Son kez işaretle
        }


        private void CheckDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(NetsisConnectionPool));
            }
        }

        // Finalizer (Güvenlik ağı olarak kalabilir ama Dispose'a güvenilmeli)
        ~NetsisConnectionPool()
        {
            _logger.LogWarning("NetsisConnectionPool finalize ediliyor - Dispose() çağrılmamış olabilir!");
            Dispose(false); // Sadece yönetilmeyen kaynaklar için (varsa), ama bizim durumumuzda Dispose(true) daha mantıklı olabilir COM için? Dikkatli olunmalı.
                            // Genellikle finalizer'dan yönetilen nesnelere (logger gibi) erişmek risklidir.
                            // Bu yüzden Dispose(false) çağrısı daha standarttır. COM temizliği için Dispose(true) çağrılmalı.
        }

        // --- Yardımcı Metodlar (Test/Debug için) ---
        /// <summary>
        /// Havuz bütünlüğünü kontrol eder (Dikkat: Anlık durumdur, thread'ler arası durumu tam yansıtmayabilir)
        /// </summary>
        public void VerifyPoolIntegrity()
        {
            int currentTotal = Volatile.Read(ref _totalConnections);
            int currentActive = Volatile.Read(ref _activeConnections);
            int inBagCount = _connections.Count;
            int semaphoreCount = _poolSemaphore.CurrentCount; // Boş slot sayısı

            // Basit kontroller
            if (currentActive < 0 || currentTotal < 0 || currentTotal > _maxPoolSize)
            {
                _logger.LogWarning($"Havuz bütünlüğü sorunu (Sayaçlar): Aktif={currentActive}, Toplam={currentTotal}, Max={_maxPoolSize}");
            }
            // Not: (inBagCount + currentActive) her zaman currentTotal'a eşit olmayabilir, çünkü
            // bir thread bağlantıyı bag'den almış ama henüz active sayacını artırmamış olabilir.
            // Veya tam tersi release durumunda. Bu nedenle bu kontrol yanıltıcı olabilir.

            // Semaphore kontrolü daha anlamlı olabilir: Max - BoşSlot = KullanılanSlot
            int usedSlots = _maxPoolSize - semaphoreCount;
            // Kullanılan slot sayısı, aktif + (oluşturulmakta olanlar) + (havuzda olup henüz alınmamışlar?) - bu da tam doğru değil.
            // En iyi kontrol: Aktif + Havuzdaki <= Toplam <= MaxPoolSize
            if ((inBagCount + currentActive) > currentTotal || currentTotal > _maxPoolSize)
            {
                _logger.LogWarning($"Havuz bütünlüğü sorunu (Sayılar): Aktif={currentActive}, Havuzda={inBagCount}, Toplam={currentTotal}, Max={_maxPoolSize}");
            }
            _logger.LogTrace($"Havuz Bütünlük Kontrolü: Aktif={currentActive}, Havuzda={inBagCount}, Toplam={currentTotal}, Max={_maxPoolSize}, Semaphore Boş={semaphoreCount}");

        }
    }

    public class NetsisConnection : IDisposable
    {
        public Kernel Kernel { get; private set; }
        public Sirket Company { get; private set; }
        public NetRS NetRS { get; private set; }
        public bool IsDisposed { get; private set; } // _disposed yerine public property

        private readonly ILogger _logger; // Gelen logger

        // Logger olmadan da çalışabilmesi için
        public NetsisConnection(Kernel kernel, Sirket company, NetRS netRS, ILogger logger = null)
        {
            Kernel = kernel ?? throw new ArgumentNullException(nameof(kernel));
            Company = company ?? throw new ArgumentNullException(nameof(company));
            NetRS = netRS ?? throw new ArgumentNullException(nameof(netRS));
            _logger = logger ?? Microsoft.Extensions.Logging.Abstractions.NullLogger.Instance;
            IsDisposed = false;
        }

        /// <summary>
        /// Bağlantıyı doğrular ve açmaya çalışır.
        /// </summary>
        /// <returns>Başarılı ise true, değilse false.</returns>
        public bool ValidateAndOpen(ILogger callerLogger = null) // Opsiyonel logger
        {
            var effectiveLogger = callerLogger ?? _logger; // Öncelik çağırandan gelen logger
            if (IsDisposed)
            {
                effectiveLogger.LogWarning("ValidateAndOpen: Bağlantı zaten dispose edilmiş.");
                return false;
            }

            try
            {
                // NetRS.Ac genellikle basit bir sorgu ile bağlantının canlı olup olmadığını kontrol eder.
                // Eğer zaten açıksa tekrar açmak sorun yaratıyor mu? Netsis dokümantasyonuna bakılmalı.
                // Genellikle idempotent olmalı (tekrar çağırmak sorun olmamalı) veya açık olup olmadığını kontrol etmeli.
                // Güvenlik için önce kapatmayı deneyebiliriz ama bu performans kaybı olabilir.
                // Şimdilik doğrudan Ac deniyoruz.
                if (NetRS.Ac("SELECT 1")) // Basit, hızlı bir sorgu
                {
                    effectiveLogger.LogTrace("ValidateAndOpen: Bağlantı başarıyla açıldı/doğrulandı.");
                    return true;
                }
                else
                {
                    // Ac false dönerse (bazı API'lerde olabilir)
                    effectiveLogger.LogWarning("ValidateAndOpen: NetRS.Ac() false döndü.");
                    return false;
                }
            }
            catch (COMException comEx)
            {
                effectiveLogger.LogError(comEx, "ValidateAndOpen: COMException sırasında hata. HResult: {HResult}", comEx.HResult);
                // Belirli HResult kodlarına göre özel işlem yapılabilir (örn. bağlantı kopmuş)
                return false;
            }
            catch (Exception ex)
            {
                effectiveLogger.LogError(ex, "ValidateAndOpen: Genel hata.");
                return false;
            }
        }


        public void Dispose()
        {
            if (IsDisposed) return;
            IsDisposed = true; // Önce işaretle

            _logger.LogTrace("NetsisConnection Dispose ediliyor...");

            // COM nesnelerini serbest bırakma sırası önemli olabilir.
            // Genellikle NetRS -> Sirket -> Kernel sırası mantıklıdır.
            try
            {
                // NetRS'yi temizle
                if (NetRS != null)
                {
                    try
                    {
                        // Kapatmayı denemek iyi olabilir ama hata verebilir, FinalReleaseComObject yine de çağrılmalı.
                        NetRS.Kapat();
                        _logger.LogTrace("NetRS kapatıldı.");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Dispose sırasında NetRS.Kapat hatası.");
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(NetRS);
                        NetRS = null; // Referansı kaldır
                        _logger.LogTrace("NetRS COM nesnesi serbest bırakıldı.");
                    }
                }

                // Sirket'i temizle
                if (Company != null)
                {
                    // Sirket nesnesinin özel bir kapatma metodu var mı? Kontrol edilmeli.
                    Marshal.FinalReleaseComObject(Company);
                    Company = null; // Referansı kaldır
                    _logger.LogTrace("Sirket COM nesnesi serbest bırakıldı.");
                }

                // Kernel'i temizle
                if (Kernel != null)
                {
                    try
                    {
                        Kernel.FreeNetsisLibrary(); // Önce kütüphaneyi serbest bırak
                        _logger.LogTrace("Kernel.FreeNetsisLibrary() çağrıldı.");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Dispose sırasında Kernel.FreeNetsisLibrary hatası.");
                    }
                    finally
                    {
                        Marshal.FinalReleaseComObject(Kernel);
                        Kernel = null; // Referansı kaldır
                        _logger.LogTrace("Kernel COM nesnesi serbest bırakıldı.");
                    }
                }
                _logger.LogInformation("NetsisConnection başarıyla dispose edildi.");
            }
            catch (Exception ex)
            {
                // Dispose içinde hata olması kötü, ama yutulmamalı en azından loglanmalı.
                _logger.LogCritical(ex, "NetsisConnection Dispose işlemi sırasında KRİTİK HATA oluştu!");
                // Bu durumda COM nesneleri sızabilir!
            }
            finally
            {
                GC.SuppressFinalize(this); // Finalizer'ı engelle
            }
        }

        // Finalizer yine de bir güvenlik ağı olarak kalabilir.
        ~NetsisConnection()
        {
            if (!IsDisposed)
            {
                // Finalizer thread'inden loglama riskli olabilir. Trace daha güvenli.
                Trace.WriteLine("UYARI: NetsisConnection nesnesi finalize ediliyor. Dispose() çağrılmamış!", "NetsisConnectionFinalizer");
                // Finalizer'dan COM nesnelerini temizlemek genellikle önerilmez ama başka çare yoksa denenebilir.
                // Ancak bu thread apartment state (STA/MTA) sorunlarına yol açabilir.
                // Dispose(true) çağırmak yerine doğrudan COM release denenebilir ama riskli.
                Dispose(); // Yönetilen/Yönetilmeyen temizlik için çağırıyoruz yine de.
            }
        }
    }
}