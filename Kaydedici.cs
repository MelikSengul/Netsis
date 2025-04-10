using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Concurrent;
using System.Diagnostics; // Debug.WriteLine için
using System.Text;
using System.Windows.Controls; // <<<=== WPF TextBox
using System.Windows.Threading; // <<<=== WPF Dispatcher

// --- WPF UYUMLU LOGGER IMPLEMENTASYONU ---

// 1. WPF Logger
public class KaydediciLogger : ILogger
{
    private readonly string _categoryName;
    private readonly LogLevel _minLevel;
    private readonly TextBox _textBox; // WPF TextBox
    private readonly int _maxTextLength; // Satır yerine karakter limiti daha kolay yönetilebilir
    private const int DefaultMaxTextLength = 50000; // Örnek limit (karakter sayısı)

    public KaydediciLogger(string categoryName, LogLevel minLevel, TextBox textBox, int maxTextLength = DefaultMaxTextLength)
    {
        _categoryName = categoryName;
        _minLevel = minLevel;
        _textBox = textBox ?? throw new ArgumentNullException(nameof(textBox));
        _maxTextLength = maxTextLength > 1000 ? maxTextLength : DefaultMaxTextLength; // Minimum bir limit belirle
    }

    // Scope'ları bu örnekte desteklemiyoruz
    public IDisposable BeginScope<TState>(TState state) => KaydediciNullScope.Instance;

    public bool IsEnabled(LogLevel logLevel) => logLevel >= _minLevel;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
    {
        if (!IsEnabled(logLevel) || formatter == null) return;

        var message = formatter(state, exception);
        if (string.IsNullOrEmpty(message) && exception == null) return;

        // Log mesajını formatla
        var logBuilder = new StringBuilder();
        logBuilder.Append($"{DateTime.Now:HH:mm:ss.fff}");
        logBuilder.Append($" [{logLevel.ToString().ToUpperInvariant().Substring(0, 3)}]");
        logBuilder.Append($" {_categoryName}:");
        logBuilder.Append($" {message}");
        if (exception != null) logBuilder.AppendLine().Append($"   >>> {exception.GetType().Name}: {exception.Message}");

        string logEntry = logBuilder.ToString() + Environment.NewLine; // Sonuna NewLine ekle

        // WPF: UI thread'inde çalıştırmak için Dispatcher kullan
        Action updateAction = () =>
        {
            try
            {
                // Maksimum karakter limitini kontrol et
                if (_textBox.Text.Length > _maxTextLength)
                {
                    // Baştan bir kısmını sil (yaklaşık %20'sini koru gibi)
                    int trimIndex = _textBox.Text.Length - (int)(_maxTextLength * 0.8);
                    if (trimIndex > 0)
                    {
                        // Daha verimli: Doğrudan Text manipülasyonu yerine Replace kullanmak
                        // veya Selection ile silmek daha iyi olabilir ama bu daha basit.
                        _textBox.Text = _textBox.Text.Substring(trimIndex);
                    }
                    else // Beklenmedik durum
                    {
                        _textBox.Text = _textBox.Text.Substring(_textBox.Text.Length / 2); // Ortadan kes
                    }

                }

                _textBox.AppendText(logEntry); // Metni ekle
                _textBox.ScrollToEnd();        // Sona kaydır
            }
            catch (Exception ex)
            {
                // TextBox'a yazarken hata olursa Debug'a yazdır
                Debug.WriteLine($"KaydediciLogger: TextBox güncelleme hatası: {ex}");
            }
        };

        try
        {
            // Dispatcher üzerinden UI thread'ine eriş
            if (_textBox.Dispatcher.CheckAccess())
            {
                updateAction(); // Zaten UI thread'indeyiz
            }
            else
            {
                // Değilsek, UI thread'ine gönder (BeginInvoke daha iyi olabilir)
                _textBox.Dispatcher.BeginInvoke(updateAction);
            }
        }
        catch (Exception ex) // Dispatcher erişiminde genel hata
        {
            Debug.WriteLine($"KaydediciLogger: Dispatcher erişim hatası: {ex}");
        }
    }
}

// 2. WPF Provider
public class KaydediciProvider : ILoggerProvider
{
    private readonly LogLevel _minLevel;
    private readonly TextBox _textBox; // WPF TextBox
    private readonly int _maxTextLength;
    private readonly ConcurrentDictionary<string, KaydediciLogger> _loggers = new ConcurrentDictionary<string, KaydediciLogger>();

    public KaydediciProvider(TextBox textBox, LogLevel minLevel, int maxTextLength)
    {
        _textBox = textBox ?? throw new ArgumentNullException(nameof(textBox));
        _minLevel = minLevel;
        _maxTextLength = maxTextLength;
    }

    public ILogger CreateLogger(string categoryName) =>
        _loggers.GetOrAdd(categoryName, name => new KaydediciLogger(name, _minLevel, _textBox, _maxTextLength));

    public void Dispose() => _loggers.Clear();
}

// 3. WPF Uzantı Metodu
public static class KaydediciExtensions
{
    public static ILoggingBuilder AddKaydediciLogger(
        this ILoggingBuilder builder,
        Func<IServiceProvider, TextBox> textBoxProvider,
        LogLevel minLevel = LogLevel.Information)
    {
        builder.Services.AddSingleton<ILoggerProvider>(provider =>
            new KaydediciProvider(textBoxProvider(provider), minLevel));
        return builder;
    }
}

// 4. Boş Kapsam (Scope)
internal sealed class KaydediciNullScope : IDisposable
{
    public static KaydediciNullScope Instance { get; } = new KaydediciNullScope();
    public void Dispose() { }
}