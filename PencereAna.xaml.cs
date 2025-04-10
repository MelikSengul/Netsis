using System;
using System.Data.SqlClient;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using NetAi;
using NetOpenX50;
using CommonQuery;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;

namespace NetAi
{
    /// <summary>
    /// Ana pencere sınıfı - Netsis ve Vepos entegrasyonunu yönetir
    /// </summary>
    public partial class PencereAna : System.Windows.Window
    {

        // Bağlantı nesneleri
        private NetsisConnectionPool _netsisConnectionPool;

        // Veritabanı bağlantı bilgileri
        string netsis_sunucu, netsis_vt_adi, netsis_vt_kull_adi, netsis_vt_kull_sifre, netsis_kull_adi, netsis_kull_sifre;
        int netsis_isletme_kodu, netsis_sube_kodu;
        string vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola;
        int bekleme_suresi = 0;

        // Entegrasyon ayarları
        string netsis_entegrasyon_fis_no_ilk_farf, netsis_entegrasyon_irs_no_ilk_farf,
               netsis_entegrasyon_nakit_kasa_kodu, netsis_entegrasyon_depo_kodu, banka_sart;

        bool cikis_yapildi = false;

        // İşlem durum takibi
        private bool isProcessing = false;
        private object processingLock = new object();

        // Log buffer
        private readonly Queue<(string mesaj, bool hataMi)> logBuffer = new Queue<(string, bool)>();
        private const int MaxBufferSize = 250;


        // Log Yeni
        private PencereLog pencereLog;
        private readonly ILogger<PencereAna> _logger;
        private readonly ILoggerFactory _loggerFactory;

        // Constructor: Gerekli bağımlılıkları (log penceresi, logger, havuz vs.) alır
        public PencereAna()
        {
            InitializeComponent();

            this.Loaded += PencereAna_Loaded;

            // ILogger'ı DI ile alıyoruz
            _logger = App.ServiceProvider.GetRequiredService<ILogger<PencereAna>>();

            // Kullanım örneği
            _logger.LogInformation("PencereAna başlatıldı.");

            // Logger'a özel provider ekle
            _loggerFactory.AddProvider(new KaydediciProvider(logPencere.LogTextBox, LogLevel.Information, 50000));

            try
            {
                throw new Exception("Test hatası");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Başlangıçta hata oluştu");
            }
        }

        private async void PencereAna_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Adım: Ayarları yükle
                UpdateProgressWithStep("Ayarlar yükleniyor", 1, 4);
                await Task.Run(() => ini_ayarlarini_oku());
                await Task.Delay(500);

                // Bağlantı havuzunu oluştur (maksimum 1 bağlantı)
                _netsisConnectionPool = new NetsisConnectionPool(
                    netsis_sunucu,
                    netsis_vt_adi,
                    netsis_vt_kull_adi,
                    netsis_vt_kull_sifre,
                    netsis_kull_adi,
                    netsis_kull_sifre,
                    netsis_sube_kodu,
                    maxPoolSize: 1);

                // 2. Adım: Netsis bağlantısı testi
                UpdateProgressWithStep("Netsis bağlantısı test ediliyor", 2, 4);
                bool netsisConnected = await TestNetsisConnectionAsync();
                await Task.Delay(1500);

                // 3. Adım: Vepos bağlantısı
                UpdateProgressWithStep("Vepos bağlantısı test ediliyor", 3, 4);
                bool veposConnected = await TestConnectionAsync($"Server={vepos_sunucu};Database={vepos_database};User Id={vepos_kullanici};Password={vepos_parola};", "Vepos");
                await Task.Delay(1000);

                // 4. Adım: Sonuç
                if (veposConnected && netsisConnected)
                {
                    UpdateProgressWithStep("Netsis-Vepos bağlantısı başarılı!", 4, 4);
                    await Task.Delay(2000);

                    //// 5. Adım: Aktarılmayı bekleyenleri başlat (en son)
                    //UpdateProgressWithStep("Aktarım bekleyenler kontrolü başlatılıyor", 5, 5);
                    //AktarmaBekleyenleriBaslat(); // Bu metod zaten async çalışıyor
                }
                else
                {
                    UpdateProgressBar("Bağlantı hatası! Aktarım kontrolü başlatılmadı");
                    _logger.LogError("Başlangıç", "Bağlantı testleri başarısız - Aktarım kontrolü çalıştırılmadı");
                }
            }
            catch (Exception ex)
            {
                UpdateProgressBar($"Hata: {ex.Message}");
                _logger.LogError("Başlangıç Hatası", ex.Message);
            }
        }

        #region Bağlantı Yönetimi

        /// <summary>
        /// Netsis bağlantısını başlatır ve kullanım sonrası otomatik temizler
        /// </summary>


        // Ana Pencere kapanırken



        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            // Temizlik işlemleri (Havuz vb.) Application Exit'te daha uygun olabilir
            _netsisConnectionPool?.Dispose();
            (_loggerFactory as IDisposable)?.Dispose(); // Factory dispose edilebilir ise et
        }

        #endregion

        #region UI Güncelleme Metodları

        /// <summary>
        /// ProgressBar'ı başlangıç değerleriyle günceller
        /// </summary>
        private void UpdateProgressBar(string message, int maxValue = 100)
        {
            Dispatcher.Invoke(() =>
            {
                PB1.Maximum = maxValue;
                PB1.Value = 0;
                var progressText = PB1.Template.FindName("ProgressText", PB1) as TextBlock;
                if (progressText != null)
                {
                    progressText.Text = message;
                }
            });
        }

        private void UpdateProgressWithStep(string message, int currentStep, int totalSteps = 5) // Changed to 5
        {
            Dispatcher.Invoke(() =>
            {
                // ProgressBar değerini hesapla (0-100 arası)
                double progressValue = (currentStep / (double)totalSteps) * 100;

                PB1.Value = progressValue;

                var progressText = PB1.Template.FindName("ProgressText", PB1) as TextBlock;
                if (progressText != null)
                {
                    progressText.Text = $"{message} (%{progressValue:0})";
                }
            });
        }

        /// <summary>
        /// ProgressBar mesajını günceller
        /// </summary>
        private void UpdateProgressBarMessage(string message, bool isWarning = false, bool isError = false)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    var progressText = PB1.Template.FindName("ProgressText", PB1) as TextBlock;
                    if (progressText != null)
                    {
                        progressText.Text = message;
                        progressText.Foreground = isError ? Brushes.Red :
                                              isWarning ? Brushes.Orange : Brushes.Black;
                    }
                });

                // DÜZELTİLMİŞ LOGLAMA
                if (isError)
                {
                    _logger.LogError("ProgressBar Hata", message);
                }
                else if (isWarning)
                {
                    _logger.LogError("ProgressBar Uyarı", message);
                }
                else if (message.IndexOf("hata", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        message.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _logger.LogInformation("ProgressBar Oto-Tespit", message);
                }
                else
                {
                    // Bilgi mesajlarında başlık kullanmıyoruz
                    _logger.LogInformation("", message);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Sistem Hatası", $"ProgressBar güncelleme hatası: {ex.Message}");
            }
        }

        /// <summary>
        /// ProgressBar değerini artırır
        /// </summary>
        private void UpdateProgressBarValue(int increment)
        {
            Dispatcher.Invoke(() =>
            {
                PB1.Value += increment;
            });
        }

        #endregion

        #region Log Yönetimi

        /// <summary>
        /// Log mesajı ekler
        /// </summary>
        private readonly object _lockObject = new object();

        //public void LogEkle(string baslik, string detay, LogSeviye seviye = LogSeviye.Bilgi)
        //{
        //    bool hataMi = (seviye == LogSeviye.Hata || seviye == LogSeviye.Kritik);

        //    // Log seviyelerini sabit genişlikte yap (8 karakter)
        //    string seviyeStr = $"[{seviye}]".PadRight(8);

        //    // Başlık kısmını formatla
        //    string baslikStr = string.IsNullOrEmpty(baslik) ? string.Empty.PadRight(2) : $"[{baslik}]";

        //    string tamMesaj = $"{seviyeStr} {baslikStr} {detay}";

        //    try
        //    {
        //        // UI işlemleri için dispatcher kontrolü
        //        Action logAction = () =>
        //        {
        //            if (hataMi)
        //            {
        //                MessageBox.Show(tamMesaj, "Hata!", MessageBoxButton.OK, MessageBoxImage.Error);
        //            }

        //            lock (_lockObject)
        //            {
        //                if (logBuffer.Count >= MaxBufferSize)
        //                {
        //                    logBuffer.Dequeue();
        //                }
        //                logBuffer.Enqueue((tamMesaj, hataMi));
        //            }

        //            if (pencereLog != null && pencereLog.IsVisible)
        //            {
        //                FlushLogBuffer();
        //            }
        //        };

        //        if (!Dispatcher.CheckAccess())
        //        {
        //            Dispatcher.BeginInvoke(logAction);
        //        }
        //        else
        //        {
        //            logAction();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        try
        //        {
        //            System.Diagnostics.Debug.WriteLine($"LogEkle hatası: {ex}");
        //            lock (_lockObject)
        //            {
        //                logBuffer.Enqueue(($"[LOGGER HATASI] {ex.Message}", true));
        //            }
        //        }
        //        catch { /* Son çare */ }
        //    }
        //}

        //// Orijinal bool parametreli versiyonunuzu koruyun (geriye uyumluluk için)
        //public void LogEkle(string mesaj, bool hataMi = false)
        //{
        //    LogEkle(hataMi ? "HATA" : "", mesaj, hataMi ? LogSeviye.Hata : LogSeviye.Bilgi);
        //}

        //// Diğer overload'lar
        //public void LogEkle(string mesaj)
        //{
        //    LogEkle(mesaj, false);
        //}

        //public void LogEkle(string baslik, string detay, bool hataMi)
        //{
        //    LogEkle(baslik, detay, hataMi ? LogSeviye.Hata : LogSeviye.Bilgi);
        //}

        //// Log seviyeleri için enum
        //public enum LogSeviye
        //{
        //    Bilgi,
        //    Uyari,
        //    Hata,
        //    Kritik
        //}

        //// 1. Sadece mesaj (otomatik Bilgi seviyesi)
        //LogEkle("İşlem başladı");

        //// 2. Mesaj + bool hata durumu
        //LogEkle("Hata oluştu", true);

        //// 3. Başlık + detay + LogSeviye
        //LogEkle("Veritabanı", "Bağlantı kuruldu", LogSeviye.Bilgi);
        //LogEkle("Yetki", "Erişim reddedildi", LogSeviye.Uyari);
        //LogEkle("Kayıt", "Veri bulunamadı", LogSeviye.Hata);

        /// <summary>
        /// Log buffer'ını temizler
        /// </summary>
        //private void FlushLogBuffer()
        //{
        //    while (logBuffer.Count > 0)
        //    {
        //        var (mesaj, hataMi) = logBuffer.Dequeue();
        //        pencereLog.LogEkle(mesaj, hataMi);
        //    }
        //}

        /// <summary>
        /// Log penceresini gösterir
        /// </summary>
        private void LogPencereGoster()
        {
            pencereLog = App.ServiceProvider.GetRequiredService<PencereLog>(); // DI'dan al

            if (!pencereLog.IsVisible)
            {
                pencereLog.Show();
            }
            else
            {
                pencereLog.Activate();
                pencereLog.Focus();
            }
        }

        /// <summary>
        /// Log penceresini kapatır
        /// </summary>
        private void LogPencereKapat()
        {
            if (pencereLog != null)
            {
                pencereLog.Hide();
            }
        }

        #endregion

        #region Buton Click Eventleri

        /// <summary>
        /// Cari dağıtım sırası butonu click eventi
        /// </summary>
        private void Button_click_caridagitimsira(object sender, RoutedEventArgs e)
        {
            // Cari dağıtım sırası işlemleri burada yapılacak
            _logger.LogInformation("Cari dağıtım sırası PASİF!!! // Geliştirme aşamasında");
            UpdateProgressBarMessage("Cari dağıtım sırası PASİF!!! // Geliştirme aşamasında");
        }

        /// <summary>
        /// Cari muhasebe kodu butonu click eventi
        /// </summary>
        private void Button_click_carimuhasebekodu(object sender, RoutedEventArgs e)
        {
            // Cari muhasebe kodu işlemleri burada yapılacak
            _logger.LogInformation("Cari muhasebe kodu PASİF!!! // Geliştirme aşamasında");
            UpdateProgressBarMessage("Cari muhasebe kodu PASİF!!! // Geliştirme aşamasında");
        }

        /// <summary>
        /// Cas terazi dosya oluştur butonu click eventi
        /// </summary>
        private async void Button_click_casterazidosyaolustur(object sender, RoutedEventArgs e)
        {
            if (isProcessing)
            {
                UpdateProgressBarMessage("Başka bir işlem zaten devam ediyor, lütfen bekleyin...");
                return;
            }

            lock (processingLock)
            {
                if (isProcessing) return;
                isProcessing = true;
            }

            Button currentButton = (Button)sender;
            currentButton.IsEnabled = false;
            SetButtonsEnabled(false);

            try
            {
                UpdateProgressBarMessage("Cas terazi dosyası oluşturma işlemi başlatılıyor...");
                PB1.Value = 0;
                PB1.Maximum = 100;

                await Task.Run(async () =>
                {
                    try
                    {
                        // 1. Veritabanından veri çekme
                        DataTable dt;
                        using (Dal db_vepos = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
                        {
                            _logger.LogInformation("Veritabanına bağlanılıyor...");
                            await Task.Delay(200);
                            dt = db_vepos.GetRecordsSp("sp_terazi_dosyasi_urunleri_cas", null);
                            UpdateProgressBarValue(10);
                            _logger.LogInformation($"{dt.Rows.Count} adet ürün bilgisi alındı");
                            await Task.Delay(300);
                        }

                        // 2. Verileri filtreleme
                        DataTable filteredDt = new DataTable();
                        filteredDt.Columns.Add("PLUNo", typeof(int));
                        filteredDt.Columns.Add("ItemCode", typeof(string));
                        filteredDt.Columns.Add("Name1", typeof(string));
                        filteredDt.Columns.Add("UPrice", typeof(string));

                        int filterProgress = 10;
                        foreach (DataRow row in dt.Rows)
                        {
                            string birimsatisfiyati1 = row["birimsatisfiyati1"].ToString() + "00";
                            filteredDt.Rows.Add(row["id"], row["kodu"], row["adi"], birimsatisfiyati1);

                            if (filteredDt.Rows.Count % 20 == 0)
                            {
                                filterProgress = 10 + (int)((filteredDt.Rows.Count / (double)dt.Rows.Count) * 15);
                                UpdateProgressBarValue(filterProgress);
                                await Task.Delay(50);
                            }
                        }
                        UpdateProgressBarValue(25);
                        await Task.Delay(200);

                        // 3. Dosya hazırlığı
                        string dosyaYolu = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "temp", "CasTeraziUrunleri.xls");
                        _logger.LogInformation($"Dosya yolu: {dosyaYolu}");
                        UpdateProgressBarValue(30);

                        string directoryPath = Path.GetDirectoryName(dosyaYolu);
                        if (!Directory.Exists(directoryPath))
                        {
                            Directory.CreateDirectory(directoryPath);
                            _logger.LogInformation($"Klasör oluşturuldu: {directoryPath}");
                            await Task.Delay(200);
                        }
                        UpdateProgressBarValue(35);

                        if (File.Exists(dosyaYolu))
                        {
                            if (IsFileLocked(dosyaYolu))
                            {
                                throw new Exception("Excel dosyası açık. Lütfen dosyayı kapatıp tekrar deneyin.");
                            }
                            File.Delete(dosyaYolu);
                            _logger.LogInformation("Mevcut dosya silindi");
                            await Task.Delay(100);
                        }
                        UpdateProgressBarValue(40);

                        // 4. Excel işlemleri
                        Excel.Application excelApp = null;
                        Excel.Workbook excelWorkbook = null;
                        Excel.Worksheet excelWorksheet = null;

                        try
                        {
                            _logger.LogInformation("Excel dosyası oluşturuluyor...");
                            excelApp = new Excel.Application();
                            excelWorkbook = excelApp.Workbooks.Add();
                            excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
                            UpdateProgressBarValue(45);
                            await Task.Delay(300);

                            // Başlık satırı
                            for (int i = 0; i < filteredDt.Columns.Count; i++)
                            {
                                excelWorksheet.Cells[1, i + 1] = filteredDt.Columns[i].ColumnName;
                                await Task.Delay(50);
                            }
                            UpdateProgressBarValue(50);

                            // Verileri yazma
                            int totalRows = filteredDt.Rows.Count;
                            for (int i = 0; i < totalRows; i++)
                            {
                                for (int j = 0; j < filteredDt.Columns.Count; j++)
                                {
                                    excelWorksheet.Cells[i + 2, j + 1] = filteredDt.Rows[i][j].ToString();
                                }

                                if (i % 10 == 0 || i == totalRows - 1)
                                {
                                    int progress = 50 + (int)((i / (double)totalRows) * 45);
                                    UpdateProgressBarValue(progress);
                                    await Task.Delay(100);

                                    if (i % 100 == 0)
                                    {
                                        _logger.LogInformation($"{i}/{totalRows} satır işlendi");
                                    }
                                }
                            }

                            // Son %5'lik kısım
                            for (int p = 95; p <= 100; p++)
                            {
                                UpdateProgressBarValue(p);
                                await Task.Delay(50);
                            }

                            // Dosyayı kaydet
                            excelWorkbook.SaveAs(dosyaYolu, Excel.XlFileFormat.xlExcel8);
                            _logger.LogInformation($"Dosya başarıyla oluşturuldu: {dosyaYolu}");
                            UpdateProgressBarMessage("Cas terazi dosyası oluşturma başarılı!");
                        }
                        finally
                        {
                            if (excelWorkbook != null)
                            {
                                excelWorkbook.Close(false);
                                Marshal.ReleaseComObject(excelWorksheet);
                                Marshal.ReleaseComObject(excelWorkbook);
                            }
                            if (excelApp != null)
                            {
                                excelApp.Quit();
                                Marshal.ReleaseComObject(excelApp);
                            }
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Hata: {ex.Message}");
                        UpdateProgressBarMessage($"Hata: {ex.Message}", true);
                        throw;
                    }
                });
            }
            catch (SqlException sqlEx)
            {
                _logger.LogError($"SQL Hatası: {sqlEx.Message}", true);
                UpdateProgressBarMessage($"SQL Hatası: {sqlEx.Message}", true);
            }
            catch (Exception ex)
            {
                _logger.LogError($"İşlem Hatası: {ex.Message}", true);
                UpdateProgressBarMessage($"İşlem Hatası: {ex.Message}", true);
            }
            finally
            {
                lock (processingLock)
                {
                    isProcessing = false;
                }
                SetButtonsEnabled(true);
                Dispatcher.Invoke(() => currentButton.IsEnabled = true);

                if (PB1.Value < 100)
                {
                    PB1.Value = 100;
                }
            }
        }

        /// <summary>
        /// Aclas terazi dosya oluştur butonu click eventi
        /// </summary>
        private async void Button_click_aclasterazidosyaolustur(object sender, RoutedEventArgs e)
        {
            if (isProcessing)
            {
                UpdateProgressBarMessage("Başka bir işlem zaten devam ediyor, lütfen bekleyin...");
                return;
            }

            lock (processingLock)
            {
                if (isProcessing) return;
                isProcessing = true;
            }

            Button currentButton = (Button)sender;
            currentButton.IsEnabled = false;
            SetButtonsEnabled(false);

            try
            {
                UpdateProgressBarMessage("Aclas terazi dosyası oluşturma işlemi başlatılıyor...");
                PB1.Value = 0;
                PB1.Maximum = 100;

                await Task.Run(() =>
                {
                    try
                    {
                        // MSSQL Server bağlantısı
                        using (Dal db_vepos = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
                        {
                            // Stored procedure'den verileri çek
                            DataTable dt = db_vepos.GetRecordsSp("sp_terazi_dosyasi_urunleri_aclas", null);
                            _logger.LogInformation($"{dt.Rows.Count} adet ürün bilgisi alındı");

                            // Metin dosyasının kaydedileceği yol
                            string dosyaYolu = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "resources", "temp", "AclasTeraziUrunleri.txt");
                            _logger.LogInformation($"Dosya yolu: {dosyaYolu}");

                            // Eğer resources/temp klasörü yoksa oluştur
                            string directoryPath = Path.GetDirectoryName(dosyaYolu);
                            if (!Directory.Exists(directoryPath))
                            {
                                Directory.CreateDirectory(directoryPath);
                                _logger.LogInformation($"Klasör oluşturuldu: {directoryPath}");
                            }

                            // Dosyanın açık olup olmadığını kontrol et
                            try
                            {
                                using (var fileStream = new FileStream(dosyaYolu, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                                {
                                    // Dosya açılabiliyorsa, işleme devam et
                                }
                            }
                            catch (IOException)
                            {
                                throw new Exception("Metin dosyası açık. Lütfen dosyayı kapatıp tekrar deneyin.");
                            }

                            // Progress güncelleme
                            int totalRows = dt.Rows.Count;
                            int processedRows = 0;

                            // Metin dosyasını oluştur ve verileri yaz
                            using (StreamWriter writer = new StreamWriter(dosyaYolu, false, Encoding.UTF8))
                            {
                                // Verileri dosyaya yaz
                                foreach (DataRow row in dt.Rows)
                                {
                                    string barkod = row["barkod"].ToString().PadRight(7); // İlk 7 karakter
                                    string adi = row["adi"].ToString().PadRight(25);      // 8'den 32'ye kadar
                                    string kodu = row["kodu"].ToString().PadRight(8);    // 32'den 40'e kadar
                                    string birimsatisfiyati1 = row["birimsatisfiyati1"].ToString();

                                    if (!string.IsNullOrEmpty(birimsatisfiyati1))
                                    {
                                        birimsatisfiyati1 += "000100001";
                                    }

                                    birimsatisfiyati1 = birimsatisfiyati1.PadLeft(12);   // 46. karakterden başlayacak

                                    // Satırı dosyaya yaz
                                    writer.WriteLine($"{barkod}\t{adi}\t{kodu}\t{birimsatisfiyati1}");

                                    // Progress güncelleme
                                    processedRows++;
                                    if (processedRows % 50 == 0 || processedRows == totalRows)
                                    {
                                        int progress = (int)((processedRows / (double)totalRows) * 100);
                                        UpdateProgressBarValue(progress);
                                        _logger.LogInformation($"{processedRows}/{totalRows} satır işlendi");
                                    }
                                }
                            }

                            _logger.LogInformation($"Dosya başarıyla oluşturuldu: {dosyaYolu}");
                            UpdateProgressBarMessage("Aclas terazi dosyası oluşturma başarılı!");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"Hata: {ex.Message}", true);
                        UpdateProgressBarMessage($"Hata: {ex.Message}", true);
                        throw;
                    }
                });
            }
            catch (SqlException sqlEx)
            {
                _logger.LogError($"SQL Hatası: {sqlEx.Message}", true);
                UpdateProgressBarMessage($"SQL Hatası: {sqlEx.Message}", true);
            }
            catch (Exception ex)
            {
                _logger.LogError($"İşlem Hatası: {ex.Message}", true);
                UpdateProgressBarMessage($"İşlem Hatası: {ex.Message}", true);
            }
            finally
            {
                lock (processingLock)
                {
                    isProcessing = false;
                }
                SetButtonsEnabled(true);
                Dispatcher.Invoke(() => currentButton.IsEnabled = true);
            }
        }

        /// <summary>
        /// Log butonu click eventi
        /// </summary>
        private void Button_click_log(object sender, RoutedEventArgs e)
        {
            if (pencereLog == null || !pencereLog.IsVisible)
            {
                LogPencereGoster();
            }
            else
            {
                LogPencereKapat();
            }
        }

        /// <summary>
        /// E-Arşiv butonu click eventi
        /// </summary>
        private void Button_click_earsiv(object sender, RoutedEventArgs e)
        {
            // E-Arşiv işlemleri burada yapılacak
            _logger.LogInformation("E-Arşiv işlemi PASİF!!! // Geliştirme aşamasında");
            UpdateProgressBarMessage("E-Arşiv işlemi PASİF!!! // Geliştirme aşamasında");
        }

        /// <summary>
        /// E-Fatura butonu click eventi
        /// </summary>
        private void Button_click_efatura(object sender, RoutedEventArgs e)
        {
            // E-Fatura işlemleri burada yapılacak
            _logger.LogInformation("E-Fatura işlemi PASİF!!! // Geliştirme aşamasında");
            UpdateProgressBarMessage("E-Fatura işlemi PASİF!!! // Geliştirme aşamasında");
        }

        /// <summary>
        /// Perakende verilerini güncelle butonu click eventi
        /// </summary>
        private async void Button_click_perakendeverileriniguncelle(object sender, RoutedEventArgs e)
        {
            if (isProcessing) return;
            UpdateProgressBarMessage("Butonlar pasif olarak ayarlandı...");
            _logger.LogInformation("Butonlar pasif olarak ayarlandı...");
            lock (processingLock)
            {
                if (isProcessing) return;
                isProcessing = true;
            }

            Button currentButton = (Button)sender;
            currentButton.IsEnabled = false;
            SetButtonsEnabled(false);

            try
            {
                UpdateProgressBarMessage("Netsis'ten Vepos'a veri entegrasyonu başlatılıyor...");

                _logger.LogInformation("============================================");
                _logger.LogInformation("Netsis'ten Vepos'a veri entegrasyonu başlatılıyor...");
                _logger.LogInformation($"Başlangıç Zamanı: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                _logger.LogInformation("============================================");

                PB1.Value = 0;
                PB1.Maximum = 100;

                await Task.Run(() =>
                {
                    try
                    {
                        Netsisten_Veposa_Entegrasyonu_Baslat();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"KRİTİK HATA: Entegrasyon sırasında beklenmeyen hata: {ex.Message}");
                        throw;
                    }
                    finally
                    {
                        lock (processingLock)
                        {
                            isProcessing = false;
                        }
                    }
                });

                UpdateProgressBarMessage("Entegrasyon başarıyla tamamlandı!");
                _logger.LogInformation("============================================");
                _logger.LogInformation("ENTEGRASYON BAŞARIYLA TAMAMLANDI");
                _logger.LogInformation($"Bitiş Zamanı: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
                _logger.LogInformation("============================================");
            }
            catch (Exception ex)
            {
                UpdateProgressBarMessage("Entegrasyon sırasında hata oluştu: " + ex.Message);
                _logger.LogError("Entegrasyon sırasında hata oluştu: " + ex.Message);

                if (ex.InnerException != null)
                {
                    _logger.LogError($"İç Hata Detayı: {ex.InnerException.Message}");
                }

                _logger.LogError($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                SetButtonsEnabled(true);
                Dispatcher.Invoke(() => currentButton.IsEnabled = true);
                _logger.LogInformation("Butonlar yeniden aktif hale getirildi");
            }
        }

        /// <summary>
        /// Netsis verilerini güncelle butonu click eventi
        /// </summary>
        private async void Button_click_netsisverileriniguncelle(object sender, RoutedEventArgs e)
        {
            if (isProcessing) return;
            UpdateProgressBarMessage("Butonlar pasif olarak ayarlandı...");
            _logger.LogInformation("Butonlar pasif olarak ayarlandı...");
            lock (processingLock)
            {
                if (isProcessing) return;
                isProcessing = true;
            }

            try
            {
                SetButtonsEnabled(false);
                PB1.Value = 0;
                PB1.Maximum = 100;

                await Task.Run(() =>
                {
                    try
                    {
                        Vepostan_Netsise_Entegrasyonu_Baslat();
                    }
                    finally
                    {
                        lock (processingLock)
                        {
                            isProcessing = false;
                        }
                    }
                });

                UpdateProgressBarMessage("İşlemler başarıyla tamamlandı!");
            }
            catch (Exception ex)
            {
                UpdateProgressBarMessage("Hata oluştu: " + ex.Message, true);
            }
            finally
            {
                SetButtonsEnabled(true);
            }
        }

        #endregion

        #region Yardımcı Metodlar

        /// <summary>
        /// INI dosyasından ayarları okur
        /// </summary>
        public void ini_ayarlarini_oku()
        { // server.ini dosyasını oku
                IniFile ini = new IniFile(System.Windows.Forms.Application.StartupPath + "\\server.ini");
                netsis_isletme_kodu = ini.IniReadValueDef("Netsis", "isletme_kodu", "1").ToInt();
                netsis_sube_kodu = ini.IniReadValueDef("Netsis", "sube_kodu", "1").ToInt();

                netsis_sunucu = ini.IniReadValueDef("Netsis", "sql_sunucu", "");
                netsis_vt_adi = ini.IniReadValueDef("Netsis", "sql_vt", "");
                netsis_vt_kull_adi = ini.IniReadValueDef("Netsis", "sql_user", "");
                netsis_vt_kull_sifre = ini.IniReadValueDef("Netsis", "sql_pass", "");

                // Diğer ayarları oku
                netsis_kull_adi = ini.IniReadValueDef("Netsis", "netsis_user", "");
                netsis_kull_sifre = ini.IniReadValueDef("Netsis", "netsis_pass", "");
                netsis_entegrasyon_fis_no_ilk_farf = ini.IniReadValueDef("Netsis", "netsis_entegrasyon_fis_no_ilk_farf", "");
                netsis_entegrasyon_irs_no_ilk_farf = ini.IniReadValueDef("Netsis", "netsis_entegrasyon_irs_no_ilk_farf", "");
                netsis_entegrasyon_nakit_kasa_kodu = ini.IniReadValueDef("Netsis", "netsis_entegrasyon_nakit_kasa_kodu", "");
                netsis_entegrasyon_depo_kodu = ini.IniReadValueDef("Netsis", "netsis_entegrasyon_depo_kodu", "");
                banka_sart = ini.IniReadValueDef("Netsis", "banka_sart", "");

                vepos_sunucu = ini.IniReadValueDef("vepos", "sunucu", "");
                vepos_database = ini.IniReadValueDef("vepos", "database", "");
                vepos_kullanici = ini.IniReadValueDef("vepos", "kullanici", "");
                vepos_parola = ini.IniReadValueDef("vepos", "parola", "");

                bekleme_suresi = ini.IniReadValueDef("vepos", "bekleme_suresi", "10").ToInt();

        }

        private DispatcherTimer syncTimer;
        private bool isSyncRunning = false;
        private readonly object syncLock = new object();

        /// <summary>
        /// Senkronizasyon butonu click eventi
        /// </summary>
        private void Button_click_sync(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;

            if (syncTimer == null)
            {
                // İlk çalıştırmada timer oluştur
                syncTimer = new DispatcherTimer();
                syncTimer.Interval = TimeSpan.FromMinutes(10);
                syncTimer.Tick += SyncTimer_Tick;

                // Buton metnini değiştir
                btn.Content = "SENKRONİZASYONU DURDUR";
                btn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD9EDB8"));

                // Hemen ilk senkronizasyonu başlat
                syncTimer.Start();
                SyncTimer_Tick(null, null);
            }
            else if (syncTimer.IsEnabled)
            {
                // Timer çalışıyorsa durdur
                syncTimer.Stop();
                btn.Content = "SENKRONİZASYONU BAŞLAT";
                btn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFFAD8D8"));
                UpdateProgressBarMessage("Senkronizasyon durduruldu");
                _logger.LogInformation("Senkronizasyon", "Kullanıcı tarafından durduruldu");
            }
            else
            {
                // Timer durmuşsa yeniden başlat
                syncTimer.Start();
                btn.Content = "SENKRONİZASYONU DURDUR";
                btn.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFD9EDB8"));
                UpdateProgressBarMessage("Senkronizasyon başlatıldı", false, false);
                _logger.LogInformation("Senkronizasyon", "Kullanıcı tarafından yeniden başlatıldı");

                // Hemen senkronizasyonu başlat
                SyncTimer_Tick(null, null);
            }
        }

        /// <summary>
        /// Timer tetiklendiğinde çalışacak senkronizasyon işlemi
        /// </summary>
        private async void SyncTimer_Tick(object sender, EventArgs e)
        {
            // Eğer önceki senkronizasyon hala çalışıyorsa atla
            lock (syncLock)
            {
                if (isSyncRunning)
                {
                    _logger.LogInformation("Senkronizasyon", "Önceki senkronizasyon hala çalışıyor, bu çalıştırma atlanıyor");
                    return;
                }
                isSyncRunning = true;
            }

            try
            {
                // UI güncelleme - işlem başladı
                UpdateProgressBarMessage("Senkronizasyon başlatılıyor...");
                _logger.LogInformation("Senkronizasyon", "Yeni senkronizasyon döngüsü başlatılıyor");

                // 1. ADIM: Netsis'ten Vepos'a veri aktarımı (her zaman çalışsın)
                await Task.Run(() =>
                {
                    UpdateProgressBarMessage("1. Adım: Netsis'ten Vepos'a veri aktarımı başlatılıyor...");
                    _logger.LogInformation("Senkronizasyon", "1. Adım: Netsis -> Vepos veri aktarımı başlatıldı");

                    try
                    {
                        Netsisten_Veposa_Entegrasyonu_Baslat();
                        _logger.LogInformation("Senkronizasyon", "1. Adım: Netsis -> Vepos veri aktarımı başarıyla tamamlandı");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Senkronizasyon Hatası", $"1. Adımda hata: {ex.Message}");
                        throw;
                    }
                });

                // 2. ADIM: Aktarılmayı bekleyenleri kontrol et ve eğer varsa aktar
                bool aktarilacakVeriVarMi = false;

                await Task.Run(() =>
                {
                    UpdateProgressBarMessage("2. Adım: Aktarılmayı bekleyen veriler kontrol ediliyor...");
                    _logger.LogInformation("Senkronizasyon", "2. Adım: Aktarılmayı bekleyenler kontrol ediliyor");

                    try
                    {
                        // Bekleyen veri olup olmadığını kontrol et
                        aktarilacakVeriVarMi = BekleyenVeriKontrolEt();

                        if (aktarilacakVeriVarMi)
                        {
                            _logger.LogInformation("Senkronizasyon", "2. Adım: Aktarılacak veri bulundu, işlem başlatılacak");
                        }
                        else
                        {
                            _logger.LogInformation("Senkronizasyon", "2. Adım: Aktarılacak veri bulunamadı, işlem atlanacak");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("Senkronizasyon Hatası", $"2. Adımda hata: {ex.Message}");
                        throw;
                    }
                });

                // 3. ADIM: Eğer aktarılacak veri varsa Vepos'tan Netsis'e aktarım yap
                if (aktarilacakVeriVarMi)
                {
                    await Task.Run(() =>
                    {
                        UpdateProgressBarMessage("3. Adım: Vepos'tan Netsis'e veri aktarımı başlatılıyor...");
                        _logger.LogInformation("Senkronizasyon", "3. Adım: Vepos -> Netsis veri aktarımı başlatıldı");

                        try
                        {
                            Vepostan_Netsise_Entegrasyonu_Baslat();
                            _logger.LogInformation("Senkronizasyon", "3. Adım: Vepos -> Netsis veri aktarımı başarıyla tamamlandı");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError("Senkronizasyon Hatası", $"3. Adımda hata: {ex.Message}");
                            throw;
                        }
                    });
                }
                else
                {
                    UpdateProgressBarMessage("3. Adım: Aktarılacak veri bulunamadı, işlem atlandı");
                    _logger.LogInformation("Senkronizasyon", "3. Adım: Aktarılacak veri olmadığı için işlem atlandı");
                }

                // Senkronizasyon tamamlandı
                UpdateProgressBarMessage($"Senkronizasyon tamamlandı. Sonraki çalışma: {DateTime.Now.AddMinutes(10):HH:mm}");
                _logger.LogInformation("Senkronizasyon", "Senkronizasyon döngüsü başarıyla tamamlandı");
            }
            catch (Exception ex)
            {
                UpdateProgressBarMessage("Senkronizasyon sırasında hata oluştu!",true, true);
                _logger.LogError("Senkronizasyon Hatası", $"Genel hata: {ex.Message}");

                // Hata durumunda bir sonraki çalışmayı 5 dakikaya ayarla
                syncTimer.Stop();
                syncTimer.Interval = TimeSpan.FromMinutes(5);
                syncTimer.Start();
                _logger.LogError("Senkronizasyon", "Hata nedeniyle bir sonraki çalışma 5 dakikaya ayarlandı");
            }
            finally
            {
                lock (syncLock)
                {
                    isSyncRunning = false;
                }

                // Hata durumunda 5 dakikaya ayarlanmışsa, bir sonraki çalışmada tekrar 10 dakikaya ayarla
                if (syncTimer.Interval.TotalMinutes == 5)
                {
                    syncTimer.Stop();
                    syncTimer.Interval = TimeSpan.FromMinutes(10);
                    syncTimer.Start();
                }
            }
        }

        /// <summary>
        /// Aktarılmayı bekleyen veri olup olmadığını kontrol eder
        /// </summary>
        private bool BekleyenVeriKontrolEt()
        {
            using (Dal db_vepos = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
            {
                // Bekleyen fatura sayısını kontrol et
                DataTable dt = db_vepos.GetRecords(
                    "SELECT COUNT(id) as adet " +
                    "FROM t_stok_hareket_m " +
                    "WHERE fisturu IN (135, 137, 138) AND (netsise_yazildi IS NULL OR netsise_yazildi=0)");

                int bekleyenFaturaSayisi = dt.Rows.Count > 0 ? Convert.ToInt32(dt.Rows[0]["adet"]) : 0;

                if (bekleyenFaturaSayisi > 0)
                {
                    _logger.LogInformation("Bekleyen Veri Kontrol", $"{bekleyenFaturaSayisi} adet bekleyen fatura bulundu");
                    return true;
                }

                _logger.LogInformation("Bekleyen Veri Kontrol", "Aktarılacak bekleyen fatura bulunamadı");
                return false;
            }
        }

        /// <summary>
        /// Veritabanı bağlantısını test eder
        /// </summary>
        private async Task<bool> TestConnectionAsync(string connectionString, string systemName)
        {
            try
            {
                string logPrefix = "[" + systemName + " Bağlantı Testi]";
                Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " bağlantısı test ediliyor..."));
                _logger.LogInformation(logPrefix + " Başlatılıyor...", false);
                _logger.LogInformation(logPrefix + " Bağlantı dizesi: " + connectionString, false);

                return await Task.Run(async () =>
                {
                    try
                    {
                        using (var testConn = new SqlConnection(connectionString))
                        {
                            // 1. Bağlantı açma
                            _logger.LogInformation(logPrefix + " Veritabanı bağlantısı açılıyor...", false);
                            await testConn.OpenAsync();
                            _logger.LogInformation(logPrefix + " Veritabanı bağlantısı başarıyla açıldı", false);

                            // 2. Test sorgusunu çalıştır
                            _logger.LogInformation(logPrefix + " Test sorgusu çalıştırılıyor...", false);
                            using (var cmd = testConn.CreateCommand())
                            {
                                cmd.CommandText = "SELECT 1";
                                var result = await cmd.ExecuteScalarAsync();

                                // 3. Sonuç kontrolü
                                if (result == null || result == DBNull.Value || Convert.ToInt32(result) != 1)
                                {
                                    _logger.LogError(logPrefix + " HATA: Test sorgusu beklenen sonucu döndürmedi!", true);
                                    Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " bağlantı testi başarısız"));
                                    return false;
                                }
                            }

                            _logger.LogInformation(logPrefix + " Test sorgusu başarılı, sonuç alındı", false);
                            Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " bağlantısı başarılı"));
                            return true;
                        }
                    }
                    catch (SqlException sqlEx)
                    {
                        string errorDetails = logPrefix + " SQL HATASI\n" +
                                            "Hata No: " + sqlEx.Number + "\n" +
                                            "Durum: " + sqlEx.State + "\n" +
                                            "Seviye: " + sqlEx.Class + "\n" +
                                            "Sunucu: " + sqlEx.Server + "\n" +
                                            "Prosedür: " + sqlEx.Procedure + "\n" +
                                            "Satır: " + sqlEx.LineNumber + "\n" +
                                            "Mesaj: " + sqlEx.Message;

                        _logger.LogError(errorDetails, true);
                        Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " bağlantı testi başarısız"));
                        return false;
                    }
                    catch (Exception ex)
                    {
                        string errorDetails = logPrefix + " SİSTEM HATASI\n" +
                                            "Mesaj: " + ex.Message + "\n" +
                                            "Stack Trace: " + ex.StackTrace;

                        _logger.LogError(errorDetails, true);
                        Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " bağlantı testi başarısız"));
                        return false;
                    }
                    finally
                    {
                        _logger.LogInformation(logPrefix + " Test süreci tamamlandı", false);
                    }
                });
            }
            catch (Exception ex)
            {
                string logPrefix = "[" + systemName + " Bağlantı Testi]";
                string errorDetails = logPrefix + " BEKLENMEYEN HATA\n" +
                                    "Mesaj: " + ex.Message + "\n" +
                                    "Stack Trace: " + ex.StackTrace;

                _logger.LogError(errorDetails, true);
                Dispatcher.Invoke(() => UpdateProgressBarMessage(systemName + " testinde sistem hatası!"));
                return false;
            }
        }

        private async Task<bool> TestNetsisConnectionAsync()
        {
            _logger.LogInformation($"Test başlamadan önce: {_netsisConnectionPool.PoolStatus}");
            const string logPrefix = "[Netsis Bağlantı Testi]";
            NetsisConnection connection = null;

            try
            {
                _netsisConnectionPool.VerifyPoolIntegrity(); // Test öncesi kontrol
                Dispatcher.Invoke(() => UpdateProgressBarMessage("Netsis bağlantısı test ediliyor..."));
                _logger.LogInformation($"{logPrefix} Başlatılıyor...", false);

                return await Task.Run(() =>
                {
                    // 1. Bağlantıyı havuzdan al
                    connection = _netsisConnectionPool.GetConnection();
                    _logger.LogInformation($"Bağlantı alındıktan sonra: {_netsisConnectionPool.PoolStatus}");

                    // Test sorgusunu çalıştır
                    object result = connection.NetRS.Ac("SELECT 1 AS TestResult");

                    // Kapsamlı null kontrolü
                    if (result == null || result == DBNull.Value)
                    {
                        _logger.LogError($"{logPrefix} Hata: Geçersiz sorgu sonucu (null)");
                        return false;
                    }

                    // Tür dönüşümü için güvenli yaklaşım
                    bool? testResult = null;

                    if (result is bool)
                    {
                        testResult = (bool)result;
                    }
                    else if (result is int)
                    {
                        testResult = Convert.ToBoolean((int)result);
                    }
                    else if (result is string)
                    {
                        bool.TryParse((string)result, out var parsed);
                        testResult = parsed;
                    }

                    // Sonuç değerlendirme
                    if (!testResult.HasValue)
                    {
                        _logger.LogError($"{logPrefix} Hata: Geçersiz sonuç türü: {result.GetType().Name}");
                        return false;
                    }

                    _logger.LogInformation($"{logPrefix} Başarılı. Sonuç: {testResult.Value}");
                    return testResult.Value;
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"{logPrefix} Kritik Hata: {ex.Message}");
                return false;
            }
            finally
            {
                if (connection != null)
                {
                    _netsisConnectionPool.ReleaseConnection(connection);
                }
                _netsisConnectionPool.VerifyPoolIntegrity(); // Test sonrası kontrol
                _logger.LogInformation($"Test tamamlandıktan sonra: {_netsisConnectionPool.PoolStatus}");
            }
        }

        /// <summary>
        /// Aktarılmayı bekleyenleri hesaplar
        /// </summary>
        public void AktarmaBekleyenleriBaslat()
        {
            // Önceki çalışan thread'i durdur
            cikis_yapildi = true;

            // Yeni thread başlat
            cikis_yapildi = false;
            Task.Run(() => Vepos_aktarilmayi_bekleyenleri_hesapla());
        }

        private async void Vepos_aktarilmayi_bekleyenleri_hesapla()
        {
            _logger.LogInformation("Aktarım Bekleyenler", "Periyodik kontrol sistemi başlatıldı");
            UpdateProgressBarMessage("Aktarım bekleyenler kontrolü aktif (5 dakikada bir)");

            while (!cikis_yapildi)
            {
                int toplamBekleyen = 0;

                try
                {
                    using (var dal = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
                    {
                        _logger.LogInformation("Aktarım Bekleyenler", "Veritabanı sorgusu başlatılıyor...");

                        DataTable dt = dal.GetRecords(
                            "SELECT 'Satış Faturası' as adi, COUNT(id) as adet " +
                            "FROM t_stok_hareket_m " +
                            "WHERE fisturu IN (135, 137, 138) AND (netsise_yazildi IS NULL OR netsise_yazildi=0)");

                        toplamBekleyen = dt.Rows.Count > 0 ? Convert.ToInt32(dt.Rows[0]["adet"]) : 0;

                        string progressMesaj = $"Bekleyen fatura: {toplamBekleyen} adet | Son kontrol: {DateTime.Now:HH:mm}";
                        UpdateProgressBarMessage(progressMesaj, toplamBekleyen > 0);

                        _logger.LogInformation("Aktarım Bekleyenler", $"Sorgu tamamlandı. Sonuç: {toplamBekleyen} bekleyen belge");

                        if (toplamBekleyen > 0)
                        {
                            _logger.LogInformation("Aktarım Bekleyenler", "Dikkat: Bekleyen belgeler var!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string hataMesaji = $"Hata: {ex.Message}";
                    _logger.LogInformation("Aktarım Bekleyenler", hataMesaji);
                    _logger.LogInformation("Aktarım Bekleyenler", $"Stack Trace: {ex.StackTrace}");

                    UpdateProgressBarMessage("Aktarım kontrolünde hata!", true);
                }

                // 5 DAKİKA BEKLE (300.000 milisaniye)
                for (int i = 0; i < 30 && !cikis_yapildi; i++)
                {
                    UpdateProgressBarMessage(
                        $"Bekleyen fatura: {toplamBekleyen} adet | " +
                        $"Sonraki kontrol için... {5 - (i / 6)} dakika kaldı");

                    await Task.Delay(30000);
                }
            }

            _logger.LogInformation("Aktarım Bekleyenler", "Periyodik kontrol sistemi durduruldu");
            UpdateProgressBarMessage("Aktarım kontrolü durduruldu");
        }

        /// <summary>
        /// Butonları aktif/pasif yapar
        /// </summary>
        private void SetButtonsEnabled(bool enabled)
        {
            Dispatcher.Invoke(() =>
            {
                button_perakendeverileriniguncelle.IsEnabled = enabled;
                button_netsisverileriniguncelle.IsEnabled = enabled;
                button_casteraziexport.IsEnabled = enabled;
                button_aclasteraziexport.IsEnabled = enabled;
            });
        }

        /// <summary>
        /// Dosyanın kilitli olup olmadığını kontrol eder
        /// </summary>
        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    stream.Close();
                }
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        #endregion

        #region Entegrasyon Metodları

        /// <summary>
        /// Netsis'ten Vepos'a veri entegrasyonunu başlatır
        /// </summary>
        private void Netsisten_Veposa_Entegrasyonu_Baslat()
        {
            using (Dal db_vepos = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
            using (Dal db_netsis = new Dal(netsis_sunucu, netsis_vt_adi, vepos_kullanici, vepos_parola, false))
            {
                try
                {
                    UpdateProgressBarMessage("Veritabanı bağlantıları başarıyla açıldı");
                    _logger.LogInformation("Veritabanı bağlantıları başarıyla açıldı");

                    // 1. Adım: Stok kartlarını güncelleme
                    UpdateProgressBarMessage("1. ADIM: Netsis'ten stok kartları çekiliyor...");
                    _logger.LogInformation("1. ADIM: Netsis'ten stok kartları çekiliyor...");
                    DataTable dt_netsis_stoklar = db_netsis.GetRecordsSp("vepos_netsis_stok_cek",
                        new SqlParameter[] {
                    new SqlParameter("@isletme", netsis_isletme_kodu),
                    new SqlParameter("@sube", netsis_sube_kodu),
                    new SqlParameter("@depo_kodu", netsis_entegrasyon_depo_kodu),
                        });

                    _logger.LogInformation($"Netsis'ten {dt_netsis_stoklar.Rows.Count} adet stok kaydı alındı");

                    // 2. Adım: Vepos'taki stokları silme
                    UpdateProgressBarMessage("2. ADIM: Vepos'taki stok kartları kontrol ediliyor...");
                    _logger.LogInformation("2. ADIM: Vepos'taki stok kartları kontrol ediliyor...");
                    DataTable dt_vepos_stoklar = db_vepos.GetRecords("select id, ozelkodu from t_stok_karti");

                    _logger.LogInformation($"Vepos'ta {dt_vepos_stoklar.Rows.Count} adet stok kaydı bulundu");

                    // 3. Adım: Barkod güncelleme
                    UpdateProgressBarMessage("3. ADIM: Netsis'ten barkod bilgileri çekiliyor...");
                    _logger.LogInformation("3. ADIM: Netsis'ten barkod bilgileri çekiliyor...");
                    DataTable dt_bar = db_netsis.GetRecordsSp("vepos_netsis_barkod_cek",
                        new SqlParameter[] {
                    new SqlParameter("@isletme", netsis_isletme_kodu),
                    new SqlParameter("@sube", netsis_sube_kodu),
                        });

                    _logger.LogInformation($"Netsis'ten {dt_bar.Rows.Count} adet barkod kaydı alındı");

                    // 4. Adım: Cari güncelleme
                    UpdateProgressBarMessage("4. ADIM: Netsis'ten cari kartlar çekiliyor...");
                    _logger.LogInformation("4. ADIM: Netsis'ten cari kartlar çekiliyor...");
                    DataTable dt_cariler = db_netsis.GetRecordsSp("vepos_netsis_cari_cek",
                        new SqlParameter[] {
                    new SqlParameter("@isletme", netsis_isletme_kodu),
                    new SqlParameter("@sube", netsis_sube_kodu),
                        });

                    _logger.LogInformation($"Netsis'ten {dt_cariler.Rows.Count} adet cari kaydı alındı");

                    // 5. Adım: Banka güncelleme
                    UpdateProgressBarMessage("5. ADIM: Netsis'ten banka bilgileri çekiliyor...");
                    _logger.LogInformation("5. ADIM: Netsis'ten banka bilgileri çekiliyor...");
                    DataTable dt_netsis_bankalar = db_netsis.GetRecordsSp("vepos_netsis_banka_kodlari",
                        new SqlParameter[] {
                    new SqlParameter("@sart", banka_sart),
                        });

                    _logger.LogInformation($"Netsis'ten {dt_netsis_bankalar.Rows.Count} adet banka kaydı alındı");

                    // 6. Adım: Vepos'taki bankaları silme
                    UpdateProgressBarMessage("6. ADIM: Vepos'taki banka bilgileri kontrol ediliyor...");
                    _logger.LogInformation("6. ADIM: Vepos'taki banka bilgileri kontrol ediliyor...");
                    DataTable dt_vepos_bankalar = db_vepos.GetRecords("select kodu from t_banka_kart");

                    _logger.LogInformation($"Vepos'ta {dt_vepos_bankalar.Rows.Count} adet banka kaydı bulundu");

                    // Toplam işlem adımlarını hesapla
                    int totalSteps = dt_netsis_stoklar.Rows.Count + dt_vepos_stoklar.Rows.Count +
                                   dt_bar.Rows.Count + dt_cariler.Rows.Count +
                                   dt_netsis_bankalar.Rows.Count + dt_vepos_bankalar.Rows.Count;

                    // ProgressBar'ı toplam adıma göre ayarla
                    UpdateProgressBar("Netsis entegrasyonu başlatılıyor...", totalSteps);
                    int currentStep = 0;

                    // 1. ADIM: Stok kartlarını güncelle
                    UpdateProgressBarMessage("1. ADIM: Stok kartları Vepos'a aktarılıyor...");
                    _logger.LogInformation("1. ADIM: Stok kartları Vepos'a aktarılıyor...");
                    foreach (DataRow dr in dt_netsis_stoklar.Rows)
                    {
                        try
                        {
                            string stokKodu = dr["STOK_KODU"].ToStr();

                            _logger.LogInformation($"Stok güncelleniyor: {stokKodu} - {dr["STOK_ADI"].ToStr()}");

                            db_vepos.ExecuteNonQuerySp("sp_netsis_ten_gelen_stogu_isle",
                                new SqlParameter[] {
        new SqlParameter("@STOK_KODU", dr["STOK_KODU"].ToStr()),
        new SqlParameter("@DEPO_KODU", dr["DEPO_KODU"].ToStr()),
        new SqlParameter("@GRUP_ISIM", dr["GRUP_ISIM"].ToStr()),
        new SqlParameter("@KOD_1", dr["KOD_1"].ToStr()),
        new SqlParameter("@BARKOD1", dr["BARKOD1"].ToStr()),
        new SqlParameter("@STOK_ADI", dr["STOK_ADI"].ToStr()),
        new SqlParameter("@MARKA", dr["MARKA"].ToStr()),
        new SqlParameter("@KDV_ORANI", dr["KDV_ORANI"].ToDbl()),
        new SqlParameter("@ALIS_FIAT1", dr["ALIS_FIAT1"].ToDbl()),
        new SqlParameter("@SATIS_FIAT1", dr["SATIS_FIAT1"].ToDbl()),
        new SqlParameter("@SATIS_FIAT2", dr["SATIS_FIAT2"].ToDbl()),
        new SqlParameter("@SATIS_FIAT3", dr["SATIS_FIAT3"].ToDbl()),
        new SqlParameter("@SATIS_FIAT4", dr["SATIS_FIAT4"].ToDbl()),
        new SqlParameter("@OLCU_BR1", dr["OLCU_BR1"].ToStr()),
        new SqlParameter("@OLCU_BR2", dr["OLCU_BR2"].ToStr()),
        new SqlParameter("@OLCU_BR3", dr["OLCU_BR3"].ToStr()),
        new SqlParameter("@PAYDA_1", dr["PAYDA_1"].ToDbl()),
        new SqlParameter("@PAYDA2", dr["PAYDA2"].ToDbl()),
        new SqlParameter("@AZAMI_STOK", dr["AZAMI_STOK"].ToDbl()),
        new SqlParameter("@MEVCUT_STOK_MIKTARI", dr["MEVCUT_STOK_MIKTARI"].ToDbl()),
        new SqlParameter("@tartim_urunu", dr["tartim_urunu"].ToStr()),
        new SqlParameter("@otomatik_tartim", dr["otomatik_tartim"].ToStr()),
        new SqlParameter("@tartim_urunu_adetli", dr["tartim_urunu_adetli"].ToStr()),
        new SqlParameter("@urun_omru", dr["urun_omru"].ToInt())
                                });

                            currentStep++;
                            UpdateProgressBarValue(1);

                            if (currentStep % 50 == 0)
                            {
                                _logger.LogInformation($"İlerleme: {currentStep}/{totalSteps} adım tamamlandı");
                            }
                        }
                        catch (Exception ex)
                        {
                            UpdateProgressBarMessage($"HATA: Stok güncellenirken hata oluştu - {dr["STOK_KODU"]} - {ex.Message}", true);
                            _logger.LogError($"HATA: Stok güncellenirken hata oluştu - {dr["STOK_KODU"]} - {ex.Message}", true);
                            throw;
                        }
                    }
                    UpdateProgressBarMessage("1. ADIM: Stok kartları başarıyla güncellendi");
                    _logger.LogInformation("1. ADIM: Stok kartları başarıyla güncellendi");

                    // 2. ADIM: Vepos'taki stokları sil
                    UpdateProgressBarMessage("2. ADIM: Vepos'taki gereksiz stoklar temizleniyor...");
                    _logger.LogInformation("2. ADIM: Vepos'taki gereksiz stoklar temizleniyor...");

                    foreach (DataRow dr in dt_vepos_stoklar.Rows)
                    {
                        try
                        {
                            string ozelKod = dr["ozelkodu"].ToStr();
                            if (dt_netsis_stoklar.Select("STOK_KODU='" + ozelKod + "'").Length == 0)
                            {
                                _logger.LogInformation($"Silinecek stok: {ozelKod}");

                                db_vepos.ExecuteNonQuery(
                                    "delete from t_stok_karti where id=" + dr["id"].ToStr() + "\n" +
                                    "delete from t_stok_birim where stok_id=" + dr["id"].ToStr() + "\n" +
                                    "delete from t_stok_resim where stok_id=" + dr["id"].ToStr() + "\n" +
                                    "delete from t_depo_mik where stok_id=" + dr["id"].ToStr()
                                );
                            }

                            currentStep++;
                            UpdateProgressBarValue(1);
                        }
                        catch (Exception ex)
                        {
                            UpdateProgressBarMessage($"HATA: Stok silinirken hata oluştu - {dr["ozelkodu"]} - {ex.Message}", true);
                            _logger.LogError($"HATA: Stok silinirken hata oluştu - {dr["ozelkodu"]} - {ex.Message}", true);
                            throw;
                        }
                    }
                    UpdateProgressBarMessage("2. ADIM: Stok temizleme işlemi tamamlandı");
                    _logger.LogInformation("2. ADIM: Stok temizleme işlemi tamamlandı");

                    // 3. ADIM: Barkod güncelle
                    UpdateProgressBarMessage("3. ADIM: Barkod bilgileri güncelleniyor...");
                    _logger.LogInformation("3. ADIM: Barkod bilgileri güncelleniyor...");
                    foreach (DataRow dr in dt_bar.Rows)
                    {
                        try
                        {
                            string stokKodu = dr["STOK_KODU"].ToStr();
                            string barkod = dr["BARKOD"].ToStr();

                            _logger.LogInformation($"Barkod güncelleniyor: {stokKodu} - {barkod}");

                            db_vepos.ExecuteNonQuery(
                                "declare @stok_id int\n" +
                                "select @stok_id=id from t_stok_karti where ozelkodu='" + dr["STOK_KODU"].ToStr() + "'\n" +
                                "if not exists(select id from t_stok_birim where stok_id=@stok_id and barkod='" + dr["BARKOD"].ToStr() + "')\n" +
                                "	insert into t_stok_birim( stok_id, barkod, birim, carpan, fiyat ) values(@stok_id, '" + dr["BARKOD"].ToStr() + "', '" + dr["birim"].ToStr() + "', " + dr["carpan"].ToSqlDbl() + ", " + dr["fiyat"].ToSqlDbl() + ")\n" +
                                "else\n" +
                                "   update t_stok_birim set birim='" + dr["birim"].ToStr() + "', carpan=" + dr["carpan"].ToSqlDbl() + ", fiyat=" + dr["fiyat"].ToSqlDbl() + " where stok_id=@stok_id and barkod='" + dr["BARKOD"].ToStr() + "'"
                            );

                            currentStep++;
                            UpdateProgressBarValue(1);
                        }
                        catch (Exception ex)
                        {
                            UpdateProgressBarMessage($"HATA: Barkod güncellenirken hata oluştu - {dr["STOK_KODU"]} - {ex.Message}", true);
                            _logger.LogError($"HATA: Barkod güncellenirken hata oluştu - {dr["STOK_KODU"]} - {ex.Message}", true);
                            throw;
                        }
                    }

                    UpdateProgressBarMessage("3. ADIM: Barkod güncelleme tamamlandı");
                    _logger.LogInformation("3. ADIM: Barkod güncelleme tamamlandı");

                    // 4. ADIM: Cari güncelle
                    UpdateProgressBarMessage("4. ADIM: Cari kartlar güncelleniyor...");
                    _logger.LogInformation("4. ADIM: Cari kartlar güncelleniyor...");
                    foreach (DataRow dr in dt_cariler.Rows)
                    {
                        try
                        {
                            string cariKod = dr["CARI_KOD"].ToStr();

                            _logger.LogInformation($"Cari güncelleniyor: {cariKod} - {dr["CARI_ISIM"].ToStr()}");

                            db_vepos.ExecuteNonQuery(
                                            "if exists(select kodu from t_cari_kart where kodu='" + dr["CARI_KOD"].ToStr() + "')\n" +
                                            "	update t_cari_kart set\n" +
                                            "		hesapadi='" + dr["CARI_ISIM"].ToStr() + "',\n" +
                                            "		adres1='" + dr["CARI_ADRES"].ToStr() + "',\n" +
                                            "		sehir='" + dr["CARI_IL"].ToStr() + "',\n" +
                                            "		semt='" + dr["CARI_ILCE"].ToStr() + "',\n" +
                                            "		vergidairesi='" + dr["VERGI_DAIRESI"].ToStr() + "',\n" +
                                            "		verginumarasi='" + dr["VERGI_NUMARASI"].ToStr() + "',\n" +
                                            "		tel1='" + dr["CARI_TEL"].ToStr() + "',\n" +
                                            "		tel2='" + dr["CARI_TEL2"].ToStr() + "',\n" +
                                            "		tel3='" + dr["CARI_TEL3"].ToStr() + "',\n" +
                                            "		ceptel1='" + dr["GSM1"].ToStr() + "',\n" +
                                            "		ceptel2='" + dr["GSM2"].ToStr() + "',\n" +
                                            "		tckimlikno='" + dr["TCKIMLIKNO"].ToStr() + "',\n" +
                                            "		hesapturu='MÜŞTERİ',\n" +
                                            "		notu='" + dr["KOSULKODU"].ToStr() + "',\n" +
                                            "		borc=" + dr["BORC"].ToSqlDbl() + ",\n" +
                                            "		alacak=" + dr["ALACAK"].ToSqlDbl() + ",\n" +
                                            "		ozelkodu='" + dr["CARI_KOD"].ToStr() + "',\n" +
                                            "		netsise_aktarildi = 1\n" +
                                            "	where kodu='" + dr["CARI_KOD"].ToStr() + "'\n" +
                                            "else begin\n" +
                                            "	insert into t_cari_kart(\n" +
                                            "       kodu, ozelkodu, hesapadi, adres1, sehir, semt, vergidairesi, verginumarasi, tel1, tel2, tel3, ceptel1, ceptel2, tckimlikno, hesapturu, notu, borc, alacak, netsise_aktarildi\n" +
                                            "	) values(\n" +
                                            "		'" + dr["CARI_KOD"].ToStr() + "', '" + dr["CARI_KOD"].ToStr() + "', '" + dr["CARI_ISIM"].ToStr() + "', '" + dr["CARI_ADRES"].ToStr() + "', '" + dr["CARI_IL"].ToStr() + "',\n" +
                                            "       '" + dr["CARI_ILCE"].ToStr() + "', '" + dr["VERGI_DAIRESI"].ToStr() + "', '" + dr["VERGI_NUMARASI"].ToStr() + "',\n" +
                                            "       '" + dr["CARI_TEL"].ToStr() + "', '" + dr["CARI_TEL2"].ToStr() + "', '" + dr["CARI_TEL3"].ToStr() + "', '" + dr["GSM1"].ToStr() + "', '" + dr["GSM2"].ToStr() + "',\n" +
                                            "       '" + dr["TCKIMLIKNO"].ToStr() + "', 'MÜŞTERİ', '" + dr["KOSULKODU"].ToStr() + "', " + dr["BORC"].ToSqlDbl() + ", " + dr["ALACAK"].ToSqlDbl() + ", 1\n" +
                                            "	)\n" +
                                            "end"
                            );

                            currentStep++;
                            UpdateProgressBarValue(1);
                        }
                        catch (Exception ex)
                        {
                            UpdateProgressBarMessage($"HATA: Cari güncellenirken hata oluştu - {dr["CARI_KOD"]} - {ex.Message}", true);
                            _logger.LogError($"HATA: Cari güncellenirken hata oluştu - {dr["CARI_KOD"]} - {ex.Message}", true);
                            throw;
                        }
                    }
                    UpdateProgressBarMessage("4. ADIM: Cari güncelleme tamamlandı");
                    _logger.LogInformation("4. ADIM: Cari güncelleme tamamlandı");

                    // 5. ADIM: Banka güncelle
                    UpdateProgressBarMessage("5. ADIM: Banka bilgileri güncelleniyor...");
                    _logger.LogInformation("5. ADIM: Banka bilgileri güncelleniyor...");

                    using (var dal = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola))
                    {
                        foreach (DataRow dr in dt_netsis_bankalar.Rows)
                        {
                            try
                            {
                                string bankaKod = dr["SOZKOD"].ToStr();
                                _logger.LogInformation($"Banka güncelleniyor: {bankaKod} - {dr["ACIKLAMA"].ToStr()}");

                                // SQL sorgusu
                                string sqlQuery = @"IF EXISTS(SELECT id FROM t_banka_kart WHERE kodu=@kodu)
                BEGIN
                    UPDATE t_banka_kart 
                    SET hesapno=@hesapno, 
                        hesapadi=@hesapadi, 
                        bankaadi=@bankaadi, 
                        subeadi=@subeadi,
                        adres=@adres,
                        semt=@semt,
                        sehir=@sehir, 
                        yetkilikisi=@yetkilikisi,
                        tel1=@tel1, 
                        tel2=@tel2, 
                        eposta=@eposta, 
                        ozel_kod=@ozel_kod 
                    WHERE kodu=@kodu
                END
                ELSE
                BEGIN
                    INSERT INTO t_banka_kart(
                        hesapno, hesapadi, bankaadi, subeadi, adres, semt, sehir, 
                        yetkilikisi, tel1, tel2, eposta, kodu, borc, alacak, 
                        pos_hesabi, ozel_kod)
                    VALUES(
                        @hesapno, @hesapadi, @bankaadi, @subeadi, @adres, @semt, @sehir,
                        @yetkilikisi, @tel1, @tel2, @eposta, @kodu, 0, 0,
                        @pos_hesabi, @ozel_kod)
                END";

                                // Parametreler
                                List<SqlParameter> parameters = new List<SqlParameter>
            {
                new SqlParameter("@hesapno", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@hesapadi", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@bankaadi", SqlDbType.VarChar) { Value = dr["ACIKLAMA"].ToStr() },
                new SqlParameter("@subeadi", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@adres", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@semt", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@sehir", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@yetkilikisi", SqlDbType.VarChar) { Value = dr["yetkili_kisi"].ToStr() },
                new SqlParameter("@tel1", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@tel2", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@eposta", SqlDbType.VarChar) { Value = "" },
                new SqlParameter("@kodu", SqlDbType.VarChar) { Value = dr["SOZKOD"].ToStr() },
                new SqlParameter("@ozel_kod", SqlDbType.VarChar) { Value = dr["SOZKOD"].ToStr() },
                new SqlParameter("@pos_hesabi", SqlDbType.Int) { Value = 1 }
            };

                                // DAL sınıfının ExecuteNonQuery metoduyla sorguyu çalıştır
                                dal.ExecuteNonQuery(sqlQuery, parameters.ToArray());

                                currentStep++;
                                UpdateProgressBarValue(1);
                            }
                            catch (Exception ex)
                            {
                                UpdateProgressBarMessage($"HATA: Banka güncellenirken hata oluştu - {dr["SOZKOD"]} - {ex.Message}", true);
                                _logger.LogError($"HATA: Banka güncellenirken hata oluştu - {dr["SOZKOD"]} - {ex.Message}", true);
                                throw;
                            }
                        }
                    }

                    UpdateProgressBarMessage("5. ADIM: Banka güncelleme tamamlandı");
                    _logger.LogInformation("5. ADIM: Banka güncelleme tamamlandı");

                    // 6. ADIM: Vepos'taki bankaları sil
                    UpdateProgressBarMessage("6. ADIM: Vepos'taki gereksiz bankalar temizleniyor...");
                    _logger.LogInformation("6. ADIM: Vepos'taki gereksiz bankalar temizleniyor...");

                    foreach (DataRow dr in dt_vepos_bankalar.Rows)
                    {
                        try
                        {
                            string bankaKod = dr["kodu"].ToStr();

                            if (dt_netsis_bankalar.Select("SOZKOD = '" + dr["kodu"].ToStr() + "' ").Length == 0)
                            {
                                _logger.LogInformation($"Silinecek banka: {bankaKod}");
                                db_vepos.ExecuteNonQuery("delete from t_banka_kart where kodu='" + dr["kodu"].ToStr() + "'");
                            }

                            currentStep++;
                            UpdateProgressBarValue(1);
                        }
                        catch (Exception ex)
                        {
                            UpdateProgressBarMessage($"HATA: Banka silinirken hata oluştu - {dr["kodu"]} - {ex.Message}", true);
                            _logger.LogError($"HATA: Banka silinirken hata oluştu - {dr["kodu"]} - {ex.Message}", true);
                            throw;
                        }
                    }
                    UpdateProgressBarMessage("6. ADIM: Banka temizleme işlemi tamamlandı");
                    _logger.LogInformation("6. ADIM: Banka temizleme işlemi tamamlandı");
                    UpdateProgressBarMessage("Tüm entegrasyon adımları başarıyla tamamlandı");
                    _logger.LogInformation("Tüm entegrasyon adımları başarıyla tamamlandı");
                }
                catch (Exception ex)
                {
                    UpdateProgressBarMessage($"KRİTİK HATA: Entegrasyon sırasında beklenmeyen hata: " + ex.Message, true);
                    _logger.LogError($"KRİTİK HATA: Entegrasyon sırasında beklenmeyen hata: {ex.Message}", true);
                    throw;
                }
                finally
                {
                    _logger.LogInformation("Veritabanı bağlantıları kapatılıyor...");
                }
            }
        }

        /// <summary>
        /// Vepos'tan Netsis'e veri entegrasyonunu başlatır (Bağlantı Havuzu Optimizasyonlu)
        /// </summary>
        private void Vepostan_Netsise_Entegrasyonu_Baslat()
        {
            using (Dal db_vepos = new Dal(vepos_sunucu, vepos_database, vepos_kullanici, vepos_parola, false))
            {
                try
                {
                    // 1. ADIM: Tek bir bağlantı ile tüm işlemleri yap
                    _netsisConnectionPool.ExecuteWithConnection(netRS =>
                    {
                        // 2. ADIM: Kernel ve Şirket bilgilerini al
                        var connection = _netsisConnectionPool.GetConnection();
                        try
                        {
                            var kernel = connection.Kernel;
                            var sirket = connection.Company;

                            // 3. ADIM: Aktarılacak faturaları getir
                            UpdateProgressBarMessage("Satış faturaları için veriler hazırlanıyor...");
                            _logger.LogInformation("NETSIS Aktarım", "Fatura listesi alınıyor");

                            DataTable dt_m = db_vepos.GetRecordsSp("sp_netsis_icin_fatura_kalem_listesi_m", null);
                            int toplamFaturaSayisi = dt_m.Rows.Count;
                            int islenenFaturaSayisi = 0;

                            _logger.LogInformation($"Toplam {toplamFaturaSayisi} fatura işlenecek", "Master liste hazır");

                            // 4. ADIM: Her faturayı işle
                            foreach (DataRow dr in dt_m.Rows)
                            {
                                string currentFisturu = dr["fisturu"].ToStr();  // Değişkenleri try dışında tanımla
                                string currentFisNo = dr["fisno"].ToStr();

                                try
                                {
                                    islenenFaturaSayisi++;
                                    string fisturu = dr["fisturu"].ToStr();
                                    string fisNo = dr["fisno"].ToStr();
                                    string belgeno = $"{dr["depo_id"]}{fisturu}{fisNo}";


                                    UpdateProgressBarMessage($"{islenenFaturaSayisi}/{toplamFaturaSayisi} - {fisturu} tipindeki {fisNo} nolu fatura işleniyor...");
                                    _logger.LogInformation("Fatura İşleniyor", $"Tip: {fisturu}, No: {fisNo}");

                                    // 5. ADIM: Fatura detaylarını al
                                    DataTable dt_fatura_kalemleri = db_vepos.GetRecordsSp(
                                        "sp_netsis_icin_fatura_kalem_listesi_d",
                                        new SqlParameter[] { new SqlParameter("@stok_har_m_id", dr["id"].ToInt()) }
                                    );

                                    // 6. ADIM: Fatura tipine göre işlem yap
                                    bool sonuc = ProcessFatura(
                                        db_vepos,
                                        netRS,
                                        kernel,
                                        sirket,
                                        dt_fatura_kalemleri,
                                        currentFisturu,
                                        dr["nk_kasa_kodu"].ToStr(),
                                        dr["kk_kasa_kod"].ToStr(),
                                        dr["soz_kod"].ToStr(),
                                        dr["vade_gun"].ToInt(),
                                        dr["cari_kodu"].ToStr(),
                                        dr["satici_adi"].ToStr(),
                                        belgeno,
                                        dr["vergino"].ToStr(),
                                        dr["tckimlikno"].ToStr(),
                                        dr["id"].ToInt()
                                    );

                                    // 7. ADIM: Sonucu işle
                                    if (sonuc)
                                    {
                                        db_vepos.ExecuteNonQuery(
                                            "UPDATE t_stok_hareket_m SET netsise_yazildi=1 WHERE id=@id",
                                            new SqlParameter("@id", dr["id"])
                                        );
                                        _logger.LogInformation("Başarılı", $"{fisNo} numaralı fatura aktarıldı");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogError("Fatura Hatası", $"Fisturu: {currentFisturu}, FisNo: {currentFisNo}\nHata: {ex.Message}\nDetay: {ex.StackTrace}",
                                        true);
                                }

                                UpdateProgressBarValue((islenenFaturaSayisi * 100) / toplamFaturaSayisi);
                            }

                            _logger.LogInformation("İşlem Tamam", $"Toplam {islenenFaturaSayisi} fatura işlendi");

                        }
                        finally
                        {
                            _netsisConnectionPool.ReleaseConnection(connection);
                        }
                    });

                    UpdateProgressBarMessage("Entegrasyon başarıyla tamamlandı");
                    UpdateProgressBarValue(100);
                }
                catch (SqlException sqlEx)
                {
                    _logger.LogError("SQL Hatası",
                        $"Hata No: {sqlEx.Number}\nMesaj: {sqlEx.Message}\nProsedür: {sqlEx.Procedure}",
                        true);
                }
                catch (Exception ex)
                {
                    _logger.LogError("Kritik Hata",
                        $"Mesaj: {ex.Message}\nStack Trace: {ex.StackTrace}",
                        true);
                }
                finally
                {
                    _logger.LogInformation("Entegrasyon Sonu", "Kaynaklar serbest bırakıldı");
                }
            }
        }

        /// <summary>
        /// Fatura tipine göre uygun işlemi yapar (Internal)
        /// </summary>
        private bool ProcessFatura(
            Dal db_vepos,
            NetRS netRS,
            Kernel kernel,
            Sirket sirket,
            DataTable dt_fatura_kalemleri,
            string fisturu,
            string nk_kasa_kodu,
            string kk_kasa_kod,
            string soz_kod,
            int vade_gun,
            string cari_kod,
            string pla_kodu,
            string belgeno,
            string vergino,
            string tckimlikno,
            int faturaId)
        {
            try
            {
                // e-Fatura kontrolü
                int adet = 0;
                if (!string.IsNullOrEmpty(vergino) || !string.IsNullOrEmpty(tckimlikno))
                {
                    adet = netRS.Ac($"SELECT COUNT(EFATVERGI) FROM EFATCARIEKR WHERE AKTIF='E' AND EFATVERGI IN ('{vergino}','{tckimlikno}')").ToInt();
                }

                // Nakit ve KK tutarları
                double nk_tut = dt_fatura_kalemleri.Rows[0]["nk_tut"].ToDbl();
                double kk_tut = dt_fatura_kalemleri.Rows[0]["kk_tut"].ToDbl();

                // Fatura tipine göre işlem
                switch (fisturu)
                {
                    case "135": // Toptan satış
                        if (adet > 0)
                            return SatisIrsaliyesiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                2, cari_kod, pla_kodu, true, netsis_entegrasyon_irs_no_ilk_farf, belgeno,
                                "", "", "0", 60, false, "", 0, 0);
                        else
                            return SatisFaturasiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                2, cari_kod, pla_kodu, true, netsis_entegrasyon_fis_no_ilk_farf, belgeno,
                                "", "", "0", 60, false, "", 0, 0);

                    case "136": // Perakende KK
                        if (adet > 0)
                            return SatisIrsaliyesiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                3, cari_kod, pla_kodu, true, netsis_entegrasyon_irs_no_ilk_farf, belgeno,
                                nk_kasa_kodu, kk_kasa_kod, "0", 1, false, soz_kod, -1 * nk_tut, -1 * kk_tut);
                        else
                            return SatisFaturasiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                3, cari_kod, pla_kodu, true, netsis_entegrasyon_fis_no_ilk_farf, belgeno,
                                nk_kasa_kodu, kk_kasa_kod, "0", 1, false, soz_kod, -1 * nk_tut, -1 * kk_tut);


                    case "137": // Perakende (Nakit)
                        _logger.LogInformation("Perakende satış faturası (Nakit) işleniyor", $"Nakit Tutar: {nk_tut}");

                         if (adet > 0)
                                return SatisIrsaliyesiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                    1, cari_kod, pla_kodu, true, netsis_entegrasyon_irs_no_ilk_farf, belgeno,
                                    nk_kasa_kodu, "", "0", 1, false, "", nk_tut, 0);

                         else
                                return SatisFaturasiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                    1, cari_kod, pla_kodu, true, netsis_entegrasyon_fis_no_ilk_farf, belgeno,
                                    nk_kasa_kodu, "", "0", 1, false, "", nk_tut, 0);


                    case "138": // Perakende (Açık)
                        _logger.LogInformation("Perakende satış faturası (Açık) işleniyor", $"Nakit Tutar: {nk_tut}, KK Tutar: {kk_tut}");


                         if (adet > 0)
                                 return SatisIrsaliyesiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                            1, cari_kod, pla_kodu, true, netsis_entegrasyon_irs_no_ilk_farf, belgeno,
                                            nk_kasa_kodu, kk_kasa_kod, "0", 1, false, soz_kod, nk_tut, kk_tut);
                          else
                                  return SatisFaturasiKaydetInternal(db_vepos, netRS, kernel, sirket, dt_fatura_kalemleri,
                                      1, cari_kod, pla_kodu, true, netsis_entegrasyon_fis_no_ilk_farf, belgeno,
                                      nk_kasa_kodu, kk_kasa_kod, "0", 1, false, soz_kod, nk_tut, kk_tut);

                    default:
                        throw new Exception($"Tanımsız fiş türü: {fisturu}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("ProcessFatura Hatası", $"FaturaID: {faturaId}, Hata: {ex.Message}", true);
                return false;
            }
        }

        /// <summary>
        /// Satış faturasını Netsis'e kaydeder
        /// </summary>
        /// <summary>
        /// Satış faturasını kaydeder (Bağlantı havuzu KULLANMAZ)
        /// </summary>
        private bool SatisFaturasiKaydetInternal(
            Dal db_vepos,
            NetRS netRS,
            Kernel kernel,
            Sirket sirket,
            DataTable dt,
            int fatura_tipi,
            string cariKod,
            string pla_kodu,
            bool kdv_dahilmi,
            string fatura_no_baslik,
            string fatura_no,
            string nk_kasa_kodu,
            string kk_kasa_kodu,
            string sube_kodu,
            int vade_gun_sayisi,
            bool otomatik_fis_al,
            string soz_kodu,
            double tah_nak_tut,
            double tah_kk_tut)
        {
            Fatura fatura = null;
            FatUst fatUst = null;
            FatKalem fatKalem = null;
            Kasa kasa = null;
            HizliTahsilat tahsil = null;
            HizliTahsilatAna tahsilAna = null;

            try
            {
                // 1. Temel ayarlar
                sirket.IntIsletmeKodu = 1;
                sirket.IntSubeKodu = 0;

                // 2. Fatura oluştur
                fatura = kernel.yeniFatura(sirket,
                    fatura_tipi == 3 ? TFaturaTip.ftAFat : TFaturaTip.ftSFat);

                // 3. Fatura numarası
                fatura_no = fatura_no.PadLeft(15 - fatura_no_baslik.Length, '0');
                string FATIRS_NO = otomatik_fis_al
                    ? fatura.YeniNumara(fatura_no_baslik)
                    : fatura_no_baslik + fatura_no;

                // 4. Fatura kontrolü
                if (netRS.Ac($"SELECT COUNT(FATIRS_NO) FROM TBLFATUIRS WHERE FATIRS_NO='{FATIRS_NO}'").ToInt() > 0)
                {
                    _logger.LogError("Fatura Uyarısı", $"{FATIRS_NO} zaten mevcut");
                    return true;
                }

                // 5. Fatura başlık bilgileri
                fatura.OtomatikCevrimYapilsin = false;
                fatura.KosulluHesapla = false;

                fatUst = fatura.Ust();
                fatUst.Sube_Kodu = 0;
                fatUst.FATIRS_NO = FATIRS_NO;
                fatUst.CariKod = cariKod;
                DateTime fistarihi = dt.Rows[0]["fistarihi"].ToDate();
                fatUst.Tarih = fistarihi;
                fatUst.ENTEGRE_TRH = fistarihi;
                fatUst.FiiliTarih = fistarihi;
                fatUst.Aciklama = "Perakende Satış";
                fatUst.FIYATTARIHI = fatUst.Tarih;
                fatUst.TIPI = fatura_tipi == 3 ? TFaturaTipi.ft_Iade : TFaturaTipi.ft_Acik;
                fatUst.ODEMETARIHI = fistarihi.AddDays(vade_gun_sayisi);

                if (!string.IsNullOrEmpty(pla_kodu))
                    fatUst.PLA_KODU = pla_kodu;

                string KOD2 = dt.Rows[0]["KOD2"].ToStr();
                if (!string.IsNullOrEmpty(KOD2))
                    fatUst.KOD2 = KOD2;

                fatUst.KDV_DAHILMI = kdv_dahilmi;

                // 6. Fatura kalemleri
                foreach (DataRow dr in dt.Rows)
                {
                    fatKalem = fatura.kalemYeni(dr["ozelkodu"].ToString());
                    fatKalem.DEPO_KODU = dr["netsis_depo_kodu"].ToInt();

                    double miktar = dr["miktar"].ToDbl();
                    double carpan = dr["carpan"].ToDbl();
                    double fiyat = dr["fiyat"].ToDbl();

                    fatKalem.STra_GCMIK = miktar * carpan;
                    fatKalem.Olcubr = dr["birim"].ToInt();
                    fatKalem.STra_NF = Math.Round(fiyat / carpan, 2);
                    fatKalem.STra_BF = Math.Round(fiyat / carpan, 2);
                    fatKalem.Isk_Flag = TFatKalemIskTipi.fkitOran;
                    fatKalem.STra_SatIsk = dr["iskonto1"].ToDbl();
                }

                // 7. Fatura hesaplama ve kayıt
                fatura.HesaplamalariYap();
                fatura.kayitYeni();

                // 8. Nakit tahsilat
                if (tah_nak_tut != 0 && fatura_tipi != 2)
                {
                    kasa = kernel.yeniKasa(sirket);
                    kasa.KsMas_Kod = nk_kasa_kodu;
                    kasa.Sube_Kodu = sube_kodu.ToInt();
                    kasa.IO = tah_nak_tut > 0 ? "G" : "C";
                    kasa.CariHareketAciklama = tah_nak_tut > 0
                        ? "Perakende Nakit Tahsilat"
                        : "Perakende Nakit Tediye";
                    kasa.Tip = "C";
                    kasa.Cari_Muh = "C";
                    kasa.Kod = cariKod;
                    kasa.Fisno = FATIRS_NO;
                    kasa.Tutar = Math.Abs(tah_nak_tut);
                    kasa.Plasiyer_Kodu = pla_kodu;
                    kasa.Tarih = fistarihi;
                    kasa.Islem(TKasaIslem.tkCariOdeme);
                }

                // 9. Kredi kartı tahsilat
                if (tah_kk_tut != 0 && fatura_tipi != 2)
                {
                    tahsilAna = kernel.YeniHizliTahsilatAna(sirket);
                    tahsilAna.IslemTarihi = fistarihi;
                    tahsilAna.KasaKod = kk_kasa_kodu;
                    tahsilAna.BelgeNo = FATIRS_NO;
                    tahsilAna.CariKod = cariKod;
                    tahsilAna.DOVTIP = 0;
                    tahsil = tahsilAna.tahsilatYeni();
                    tahsil.Aciklama = tah_kk_tut > 0
                        ? "Perakende K.Kartı Tahsilat"
                        : "Perakende K.Kartı Tediye";
                    tahsil.SozKodu = soz_kodu;
                    tahsil.Tutar = Math.Abs(tah_kk_tut);
                    tahsil.PLA_KODU = pla_kodu;
                    tahsil.KartNo = "0000000000000000";
                    tahsil.TaksitSay = 1;
                    tahsilAna.kayitYeni();
                }

                _logger.LogInformation("Fatura Başarılı", $"{FATIRS_NO} kaydedildi");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError("Fatura Hatası",
                    $"FaturaNo: {fatura_no}, Hata: {ex.Message}\n{ex.StackTrace}",
                    true);
                return false;
            }
            finally
            {
                // COM nesnelerini temizle
                if (fatKalem != null) Marshal.ReleaseComObject(fatKalem);
                if (fatUst != null) Marshal.ReleaseComObject(fatUst);
                if (fatura != null) Marshal.ReleaseComObject(fatura);
                if (kasa != null) Marshal.ReleaseComObject(kasa);
                if (tahsil != null) Marshal.ReleaseComObject(tahsil);
                if (tahsilAna != null) Marshal.ReleaseComObject(tahsilAna);
            }
        }

        /// <summary>
        /// Satış irsaliyesini kaydeder (Bağlantı havuzu KULLANMAZ)
        /// </summary>
        private bool SatisIrsaliyesiKaydetInternal(
            Dal db_vepos,
            NetRS netRS,
            Kernel kernel,
            Sirket sirket,
            DataTable dt,
            int fatura_tipi,
            string cariKod,
            string pla_kodu,
            bool kdv_dahilmi,
            string irs_no_baslik,
            string irs_no,
            string nk_kasa_kodu,
            string kk_kasa_kodu,
            string sube_kodu,
            int vade_gun_sayisi,
            bool otomatik_fis_al,
            string soz_kodu,
            double tah_nak_tut,
            double tah_kk_tut)
        {
            Fatura fatura = null;
            FatUst fatUst = null;
            FatKalem fatKalem = null;
            Kasa kasa = null;
            HizliTahsilat tahsil = null;
            HizliTahsilatAna tahsilAna = null;

            try
            {
                // 1. Temel ayarlar
                sirket.IntIsletmeKodu = 1;
                sirket.IntSubeKodu = 0;

                // 2. İrsaliye oluştur
                fatura = kernel.yeniFatura(sirket, TFaturaTip.ftSIrs);

                // 3. İrsaliye numarası
                irs_no = irs_no.PadLeft(15 - irs_no_baslik.Length, '0');
                string FATIRS_NO = otomatik_fis_al
                    ? fatura.YeniNumara(irs_no_baslik)
                    : irs_no_baslik + irs_no;

                // 4. İrsaliye kontrolü
                if (netRS.Ac($"SELECT COUNT(FATIRS_NO) FROM TBLFATUIRS WHERE FATIRS_NO='{FATIRS_NO}'").ToInt() > 0)
                {
                    _logger.LogError("İrsaliye Uyarısı", $"{FATIRS_NO} zaten mevcut");
                    return true;
                }

                // 5. İrsaliye başlık bilgileri
                fatura.OtomatikCevrimYapilsin = false;
                fatura.KosulluHesapla = false;

                fatUst = fatura.Ust();
                fatUst.Sube_Kodu = 0;
                fatUst.FATIRS_NO = FATIRS_NO;
                fatUst.CariKod = cariKod;
                DateTime fistarihi = dt.Rows[0]["fistarihi"].ToDate();
                fatUst.Tarih = fistarihi;
                fatUst.ENTEGRE_TRH = fistarihi;
                fatUst.FiiliTarih = fistarihi;
                fatUst.Aciklama = "Perakende Satış İrsaliyesi";
                fatUst.FIYATTARIHI = fatUst.Tarih;
                fatUst.TIPI = fatura_tipi == 3 ? TFaturaTipi.ft_Iade : TFaturaTipi.ft_Acik;
                fatUst.ODEMETARIHI = fistarihi.AddDays(vade_gun_sayisi);

                if (!string.IsNullOrEmpty(pla_kodu))
                    fatUst.PLA_KODU = pla_kodu;

                string KOD2 = dt.Rows[0]["KOD2"].ToStr();
                if (!string.IsNullOrEmpty(KOD2))
                    fatUst.KOD2 = KOD2;

                fatUst.KDV_DAHILMI = kdv_dahilmi;

                // 6. İrsaliye kalemleri
                foreach (DataRow dr in dt.Rows)
                {
                    fatKalem = fatura.kalemYeni(dr["ozelkodu"].ToString());
                    fatKalem.DEPO_KODU = dr["netsis_depo_kodu"].ToInt();

                    double miktar = dr["miktar"].ToDbl();
                    double carpan = dr["carpan"].ToDbl();
                    double fiyat = dr["fiyat"].ToDbl();

                    fatKalem.STra_GCMIK = miktar * carpan;
                    fatKalem.Olcubr = dr["birim"].ToInt();
                    fatKalem.STra_NF = Math.Round(fiyat / carpan, 2);
                    fatKalem.STra_BF = Math.Round(fiyat / carpan, 2);
                    fatKalem.Isk_Flag = TFatKalemIskTipi.fkitOran;
                    fatKalem.STra_SatIsk = dr["iskonto1"].ToDbl();
                }

                // 7. İrsaliye hesaplama ve kayıt
                fatura.HesaplamalariYap();
                fatura.kayitYeni();

                // 8. Nakit tahsilat (fatura ile aynı)
                if (tah_nak_tut != 0 && fatura_tipi != 2)
                {
                    kasa = kernel.yeniKasa(sirket);
                    kasa.KsMas_Kod = nk_kasa_kodu;
                    kasa.Sube_Kodu = sube_kodu.ToInt();
                    kasa.IO = tah_nak_tut > 0 ? "G" : "C";
                    kasa.CariHareketAciklama = tah_nak_tut > 0
                        ? "Perakende Nakit Tahsilat"
                        : "Perakende Nakit Tediye";
                    kasa.Tip = "C";
                    kasa.Cari_Muh = "C";
                    kasa.Kod = cariKod;
                    kasa.Fisno = FATIRS_NO;
                    kasa.Tutar = Math.Abs(tah_nak_tut);
                    kasa.Plasiyer_Kodu = pla_kodu;
                    kasa.Tarih = fistarihi;
                    kasa.Islem(TKasaIslem.tkCariOdeme);
                }

                // 9. Kredi kartı tahsilat (fatura ile aynı)
                if (tah_kk_tut != 0 && fatura_tipi != 2)
                {
                    tahsilAna = kernel.YeniHizliTahsilatAna(sirket);
                    tahsilAna.IslemTarihi = fistarihi;
                    tahsilAna.KasaKod = kk_kasa_kodu;
                    tahsilAna.BelgeNo = FATIRS_NO;
                    tahsilAna.CariKod = cariKod;
                    tahsilAna.DOVTIP = 0;
                    tahsil = tahsilAna.tahsilatYeni();
                    tahsil.Aciklama = tah_kk_tut > 0
                        ? "Perakende K.Kartı Tahsilat"
                        : "Perakende K.Kartı Tediye";
                    tahsil.SozKodu = soz_kodu;
                    tahsil.Tutar = Math.Abs(tah_kk_tut);
                    tahsil.PLA_KODU = pla_kodu;
                    tahsil.KartNo = "0000000000000000";
                    tahsil.TaksitSay = 1;
                    tahsilAna.kayitYeni();
                }

                _logger.LogInformation("İrsaliye Başarılı", $"{FATIRS_NO} kaydedildi");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError("İrsaliye Hatası",
                    $"İrsaliyeNo: {irs_no}, Hata: {ex.Message}\n{ex.StackTrace}",
                    true);
                return false;
            }
            finally
            {
                // COM nesnelerini temizle
                if (fatKalem != null) Marshal.ReleaseComObject(fatKalem);
                if (fatUst != null) Marshal.ReleaseComObject(fatUst);
                if (fatura != null) Marshal.ReleaseComObject(fatura);
                if (kasa != null) Marshal.ReleaseComObject(kasa);
                if (tahsil != null) Marshal.ReleaseComObject(tahsil);
                if (tahsilAna != null) Marshal.ReleaseComObject(tahsilAna);
            }
        }

        #endregion
    }
}