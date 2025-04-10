using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using System.Windows;
using System.Windows.Controls;

namespace NetAi
{
    public partial class App : Application
    {
        public static IServiceProvider ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var services = new ServiceCollection();

            // PencereLog'u DI'ye ekle (TextBox'ı içeriyor)
            services.AddSingleton<PencereLog>();

            // Logger servislerini ekle
            services.AddLogging(builder =>
            {
                builder.AddKaydediciLogger(provider =>
                {
                    var pencereLog = provider.GetRequiredService<PencereLog>();
                    return pencereLog.LogTextBox;
                }, LogLevel.Information);
            });

            // Ana pencereyi ekle
            services.AddSingleton<PencereAna>();

            ServiceProvider = services.BuildServiceProvider();

            var pencere = ServiceProvider.GetRequiredService<PencereAna>();
            pencere.Show();
        }
    }
}