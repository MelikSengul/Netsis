using System;
using System.Collections.Concurrent;
using System.Data;
using System.Data.SqlClient;

namespace NetAi
{
    /// <summary>
    /// Veritabanı işlemleri için geliştirilmiş Data Access Layer (DAL) sınıfı
    /// Bağlantı havuzlama, transaction yönetimi ve thread-safe işlemler destekler
    /// Varsayılan bağlantı bilgileri:
    /// Sunucu: "." (local server)
    /// Veritabanı: "VEPOS"
    /// Kullanıcı: "saydam"
    /// Şifre: "saytek"
    /// </summary>
    public sealed class Dal : IDisposable
    {
        #region Varsayılan Bağlantı Bilgileri
        private const string DEFAULT_SERVER = ".";           // Local server
        private const string DEFAULT_DATABASE = "VEPOS";     // VEPOS veritabanı
        private const string DEFAULT_USER = "saydam";         // Kullanıcı adı
        private const string DEFAULT_PASSWORD = "saytek";     // Kullanıcı şifresi
        #endregion

        #region Bağlantı Havuzu (Connection Pool) Implementasyonu

        // Bağlantı havuzunu tutacak thread-safe koleksiyon
        private static readonly ConcurrentDictionary<string, SqlConnection> _connectionPool =
                new ConcurrentDictionary<string, SqlConnection>();

        // Havuz işlemleri için kilitleme mekanizması
        private static readonly object _poolLock = new object();

        // Havuz başlatıldı mı kontrolü
        private static bool _poolInitialized = false;

        // Maksimum bağlantı sayısı (varsayılan 100)
        private static int _maxPoolSize = 100;

        // Minimum bağlantı sayısı (varsayılan 5)
        private static int _minPoolSize = 5;

        // Bağlantı zaman aşımı (saniye)
        private static int _connectionTimeout = 30;

        /// <summary>
        /// Bağlantı havuzunu başlatır ve yapılandırır
        /// </summary>
        /// <param name="maxPoolSize">Maksimum bağlantı sayısı</param>
        /// <param name="minPoolSize">Minimum bağlantı sayısı</param>
        /// <param name="connectionTimeout">Bağlantı zaman aşımı (saniye)</param>
        public static void InitializeConnectionPool(int maxPoolSize = 100, int minPoolSize = 5, int connectionTimeout = 30)
        {
            if (!_poolInitialized)
            {
                lock (_poolLock)
                {
                    if (!_poolInitialized)
                    {
                        _maxPoolSize = maxPoolSize;
                        _minPoolSize = minPoolSize;
                        _connectionTimeout = connectionTimeout;
                        _poolInitialized = true;

                        // Minimum bağlantı sayısı kadar bağlantı önceden oluşturuluyor
                        for (int i = 0; i < _minPoolSize; i++)
                        {
                            var dummyConn = CreateNewConnection(DEFAULT_SERVER, DEFAULT_DATABASE, DEFAULT_USER, DEFAULT_PASSWORD);
                            _connectionPool.TryAdd(Guid.NewGuid().ToString(), dummyConn);
                        }
                    }
                }
            }
        }

        // Yeni bağlantı oluşturma metodu
        private static SqlConnection CreateNewConnection(string server, string database, string user, string password)
        {
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = server,
                InitialCatalog = database,
                UserID = user,
                Password = password,
                Pooling = true, // Bağlantı havuzlamayı aktif et
                MaxPoolSize = _maxPoolSize,
                MinPoolSize = _minPoolSize,
                ConnectTimeout = _connectionTimeout,
                ApplicationName = "VeposNetsisIntegration",
                Enlist = false,
                MultipleActiveResultSets = true, // Aynı bağlantıda birden fazla sonuç seti
                PersistSecurityInfo = false
            };

            var connection = new SqlConnection(builder.ToString());
            connection.Open();
            return connection;
        }

        /// <summary>
        /// Bağlantı havuzunu temizler ve tüm bağlantıları kapatır
        /// Uygulama kapatılırken çağrılmalıdır
        /// </summary>
        public static void CleanupConnectionPool()
        {
            foreach (var connEntry in _connectionPool)
            {
                try
                {
                    if (connEntry.Value.State == ConnectionState.Open)
                    {
                        connEntry.Value.Close();
                    }
                    connEntry.Value.Dispose();
                }
                catch { /* Temizleme hataları göz ardı ediliyor */ }
            }
            _connectionPool.Clear();
        }

        #endregion

        #region Örnek (Instance) Üyeler

        private SqlConnection _connection; // Mevcut bağlantı
        private SqlCommand _command;      // SQL komut nesnesi
        private SqlTransaction _transaction; // Transaction nesnesi
        private readonly string _connectionKey; // Bağlantı anahtarı (sunucu+veritabanı+kullanıcı)
        private bool _disposed = false;   // Dispose durumu

        // Sadece okunabilir özellikler
        public string Server { get; }     // Sunucu adı
        public string Database { get; }   // Veritabanı adı
        public string User { get; }       // Kullanıcı adı
        public bool IsInTransaction => _transaction != null; // Transaction durumu

        /// <summary>
        /// DAL sınıfı constructor'ı - Varsayılan VEPOS bağlantısı ile oluşturur
        /// </summary>
        /// <param name="useTransaction">Transaction kullanılsın mı?</param>
        public Dal(bool useTransaction = false)
            : this(DEFAULT_SERVER, DEFAULT_DATABASE, DEFAULT_USER, DEFAULT_PASSWORD, useTransaction)
        {
        }

        /// <summary>
        /// DAL sınıfı constructor'ı - Özel bağlantı bilgileri ile oluşturur
        /// </summary>
        /// <param name="server">Sunucu adı</param>
        /// <param name="database">Veritabanı adı</param>
        /// <param name="user">Kullanıcı adı</param>
        /// <param name="password">Şifre</param>
        /// <param name="useTransaction">Transaction kullanılsın mı?</param>
        public Dal(string server, string database, string user, string password, bool useTransaction = false)
        {
            Server = server ?? throw new ArgumentNullException(nameof(server));
            Database = database ?? throw new ArgumentNullException(nameof(database));
            User = user ?? throw new ArgumentNullException(nameof(user));

            // Bağlantı anahtarını oluştur (havuzda benzersiz tanımlama için)
            _connectionKey = $"{server}|{database}|{user}";

            // Havuzdan bağlantı al veya yeni oluştur
            if (!_connectionPool.TryGetValue(_connectionKey, out _connection))
            {
                lock (_poolLock)
                {
                    if (!_connectionPool.TryGetValue(_connectionKey, out _connection))
                    {
                        _connection = CreateNewConnection(server, database, user, password);
                        _connectionPool.TryAdd(_connectionKey, _connection);
                    }
                }
            }

            _command = new SqlCommand { Connection = _connection };

            // Transaction başlatma isteği varsa
            if (useTransaction)
            {
                BeginTransaction();
            }
        }

        #endregion

        #region Transaction Yönetimi

        /// <summary>
        /// Yeni bir transaction başlatır
        /// </summary>
        /// <param name="isolationLevel">İzolasyon seviyesi (varsayılan: ReadCommitted)</param>
        public void BeginTransaction(IsolationLevel isolationLevel = IsolationLevel.ReadCommitted)
        {
            if (_transaction != null)
            {
                throw new InvalidOperationException("Zaten bir transaction işlemi devam ediyor.");
            }

            _transaction = _connection.BeginTransaction(isolationLevel);
            _command.Transaction = _transaction;
        }

        /// <summary>
        /// Aktif transaction'ı commit eder (kaydeder)
        /// </summary>
        public void CommitTransaction()
        {
            if (_transaction == null)
            {
                throw new InvalidOperationException("Commit edilecek transaction bulunamadı.");
            }

            try
            {
                _transaction.Commit();
            }
            finally
            {
                _transaction.Dispose();
                _transaction = null;
                _command.Transaction = null;
            }
        }

        /// <summary>
        /// Aktif transaction'ı rollback eder (geri alır)
        /// </summary>
        public void RollbackTransaction()
        {
            if (_transaction == null)
            {
                throw new InvalidOperationException("Rollback yapılacak transaction bulunamadı.");
            }

            try
            {
                _transaction.Rollback();
            }
            finally
            {
                _transaction.Dispose();
                _transaction = null;
                _command.Transaction = null;
            }
        }

        #endregion

        #region Veri Erişim Metodları

        /// <summary>
        /// SQL sorgusu çalıştırır ve sonuçları DataTable olarak döndürür
        /// </summary>
        /// <param name="sqlQuery">Çalıştırılacak SQL sorgusu</param>
        /// <returns>Sonuç DataTable'ı</returns>
        public DataTable GetRecords(string sqlQuery)
        {
            if (string.IsNullOrWhiteSpace(sqlQuery))
                throw new ArgumentException("SQL sorgusu boş olamaz.", nameof(sqlQuery));

            var dataTable = new DataTable();

            try
            {
                _command.CommandType = CommandType.Text;
                _command.CommandText = sqlQuery;
                _command.Parameters.Clear();

                using (var adapter = new SqlDataAdapter(_command))
                {
                    adapter.Fill(dataTable);
                }
            }
            catch (Exception ex)
            {
                throw new DataAccessException("SQL sorgusu çalıştırılırken hata oluştu.", ex);
            }

            return dataTable;
        }

        /// <summary>
        /// Stored procedure çalıştırır ve sonuçları DataTable olarak döndürür
        /// </summary>
        /// <param name="storedProcedureName">Çalıştırılacak SP adı</param>
        /// <param name="parameters">SP parametreleri</param>
        /// <returns>Sonuç DataTable'ı</returns>
        public DataTable GetRecordsSp(string storedProcedureName, params SqlParameter[] parameters)
        {
            if (string.IsNullOrWhiteSpace(storedProcedureName))
                throw new ArgumentException("Stored procedure adı boş olamaz.", nameof(storedProcedureName));

            var dataTable = new DataTable();

            try
            {
                _command.CommandType = CommandType.StoredProcedure;
                _command.CommandText = storedProcedureName;
                _command.Parameters.Clear();

                if (parameters != null && parameters.Length > 0)
                {
                    _command.Parameters.AddRange(parameters);
                }

                using (var adapter = new SqlDataAdapter(_command))
                {
                    adapter.Fill(dataTable);
                }
            }
            catch (Exception ex)
            {
                throw new DataAccessException($"'{storedProcedureName}' stored procedure çalıştırılırken hata oluştu.", ex);
            }

            return dataTable;
        }

        /// <summary>
        /// Tek bir değer döndüren SQL sorgusu çalıştırır
        /// </summary>
        /// <param name="sqlQuery">Çalıştırılacak SQL sorgusu</param>
        /// <returns>Sorgu sonucu (ilk sütunun ilk satırı)</returns>
        public object ExecuteScalar(string sqlQuery)
        {
            if (string.IsNullOrWhiteSpace(sqlQuery))
                throw new ArgumentException("SQL sorgusu boş olamaz.", nameof(sqlQuery));

            try
            {
                _command.CommandType = CommandType.Text;
                _command.CommandText = sqlQuery;
                _command.Parameters.Clear();

                return _command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw new DataAccessException("Scalar sorgu çalıştırılırken hata oluştu.", ex);
            }
        }

        /// <summary>
        /// Tek bir değer döndüren stored procedure çalıştırır
        /// </summary>
        /// <param name="storedProcedureName">Çalıştırılacak SP adı</param>
        /// <param name="parameters">SP parametreleri</param>
        /// <returns>SP sonucu (ilk sütunun ilk satırı)</returns>
        public object ExecuteScalarSp(string storedProcedureName, params SqlParameter[] parameters)
        {
            if (string.IsNullOrWhiteSpace(storedProcedureName))
                throw new ArgumentException("Stored procedure adı boş olamaz.", nameof(storedProcedureName));

            try
            {
                _command.CommandType = CommandType.StoredProcedure;
                _command.CommandText = storedProcedureName;
                _command.Parameters.Clear();

                if (parameters != null && parameters.Length > 0)
                {
                    _command.Parameters.AddRange(parameters);
                }

                return _command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                throw new DataAccessException($"'{storedProcedureName}' scalar stored procedure çalıştırılırken hata oluştu.", ex);
            }
        }

        /// <summary>
        /// INSERT, UPDATE, DELETE gibi sonuç döndürmeyen SQL sorgusu çalıştırır
        /// </summary>
        /// <param name="sqlQuery">Çalıştırılacak SQL sorgusu</param>
        /// <returns>Etkilenen satır sayısı</returns>
        public int ExecuteNonQuery(string sqlQuery, params SqlParameter[] parameters)
        {
            if (string.IsNullOrWhiteSpace(sqlQuery))
                throw new ArgumentException("SQL sorgusu boş olamaz.", nameof(sqlQuery));

            try
            {
                _command.CommandType = CommandType.Text;
                _command.CommandText = sqlQuery;
                _command.Parameters.Clear();

                if (parameters != null && parameters.Length > 0)
                {
                    _command.Parameters.AddRange(parameters);
                }

                return _command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new DataAccessException("Non-query sorgu çalıştırılırken hata oluştu.", ex);
            }
        }

        /// <summary>
        /// INSERT, UPDATE, DELETE gibi sonuç döndürmeyen stored procedure çalıştırır
        /// </summary>
        /// <param name="storedProcedureName">Çalıştırılacak SP adı</param>
        /// <param name="parameters">SP parametreleri</param>
        /// <returns>Etkilenen satır sayısı</returns>
        public int ExecuteNonQuerySp(string storedProcedureName, params SqlParameter[] parameters)
        {
            if (string.IsNullOrWhiteSpace(storedProcedureName))
                throw new ArgumentException("Stored procedure adı boş olamaz.", nameof(storedProcedureName));

            try
            {
                _command.CommandType = CommandType.StoredProcedure;
                _command.CommandText = storedProcedureName;
                _command.Parameters.Clear();

                if (parameters != null && parameters.Length > 0)
                {
                    _command.Parameters.AddRange(parameters);
                }

                return _command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new DataAccessException($"'{storedProcedureName}' non-query stored procedure çalıştırılırken hata oluştu.", ex);
            }
        }

        /// <summary>
        /// Output parametreli stored procedure çalıştırır
        /// </summary>
        /// <param name="storedProcedureName">Çalıştırılacak SP adı</param>
        /// <param name="parameters">Giriş/çıkış parametreleri</param>
        /// <returns>Parametre koleksiyonu (çıkış değerlerini içerir)</returns>
        public SqlParameterCollection ExecuteSpWithOutputParameters(string storedProcedureName, params SqlParameter[] parameters)
        {
            if (string.IsNullOrWhiteSpace(storedProcedureName))
                throw new ArgumentException("Stored procedure adı boş olamaz.", nameof(storedProcedureName));

            try
            {
                _command.CommandType = CommandType.StoredProcedure;
                _command.CommandText = storedProcedureName;
                _command.Parameters.Clear();

                if (parameters != null && parameters.Length > 0)
                {
                    foreach (var param in parameters)
                    {
                        if (param.Direction != ParameterDirection.Input)
                        {
                            param.Value = DBNull.Value; // Output parametreleri başlat
                        }
                        _command.Parameters.Add(param);
                    }
                }

                _command.ExecuteNonQuery();

                return _command.Parameters;
            }
            catch (Exception ex)
            {
                throw new DataAccessException($"'{storedProcedureName}' output parametreli stored procedure çalıştırılırken hata oluştu.", ex);
            }
        }

        #endregion

        #region IDisposable Implementasyonu

        /// <summary>
        /// Kaynakları serbest bırakır
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Yönetilen kaynakları serbest bırak
                    if (_transaction != null)
                    {
                        try
                        {
                            if (_transaction.Connection != null)
                            {
                                _transaction.Rollback();
                            }
                        }
                        catch { /* Rollback hataları göz ardı ediliyor */ }
                        _transaction.Dispose();
                    }

                    _command?.Dispose();

                    // Bağlantıyı havuzda bırak, dispose etme
                    // Gerçek temizlik CleanupConnectionPool ile yapılacak
                }

                _disposed = true;
            }
        }

        // Finalizer - sadece yönetilmeyen kaynaklar için
        ~Dal()
        {
            Dispose(false);
        }

        #endregion
    }

    /// <summary>
    /// DAL katmanından fırlatılacak özel istisna sınıfı
    /// </summary>
    public class DataAccessException : Exception
    {
        public DataAccessException(string message) : base(message) { }
        public DataAccessException(string message, Exception innerException) : base(message, innerException) { }
    }
}