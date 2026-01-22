using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Collections.Concurrent;
using System.Windows;

namespace Mooseware.DvarDeputy;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{

    public static IHost? AppHost { get; private set; }

    public App()
    {
        // Establish both a main window singleton and a concurrent queue singleton to be used for the HTTP API
        AppHost = (IHost?)Host.CreateDefaultBuilder()
        .ConfigureServices((hostContext, services) =>
        {
            services.Configure<Configuration.AppSettings>(hostContext.Configuration.GetSection("ApplicationSettings"));
            services.AddSingleton<MainWindow>();
            services.AddSingleton<ConcurrentQueue<ApiMessage>>();
        })
        .Build();
    }

    protected override async void OnStartup(StartupEventArgs e)
    {
        await AppHost!.StartAsync();

        // Launch the MainWindow
        var startupForm = AppHost.Services.GetRequiredService<MainWindow>();
        startupForm.Show();

        base.OnStartup(e);
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        // Wind down the API server...
        await AppHost!.StopAsync();
        base.OnExit(e);
    }
}
