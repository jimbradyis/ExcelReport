using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.IO;
using System.Windows;
using ExcelReport.Models;
using Microsoft.EntityFrameworkCore;
// REMOVE: using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing; // Not needed unless you're using OpenXML directly

namespace ExcelReport
{
    public partial class App : Application
    {
        public static IConfiguration Configuration { get; private set; }

        // Keep a reference to the DI host (so we can resolve services later if needed)
        private IHost _host;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 1) Build configuration
            var configBuilder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            Configuration = configBuilder.Build();

            // 2) Create Host with DI
            _host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    // 2a) Register the EF DbContext
                    var connString = Configuration.GetConnectionString("SqliteConnection");
                    services.AddDbContext<EthicsContext>(options =>
                    {
                        options.UseSqlite(connString);
                    });

                    // 2b) Register your custom service(s)
                    services.AddTransient<Services.HollingerReportService>();

                    // 2c) Register MainWindow
                    services.AddTransient<MainWindow>();
                })
                .Build();

            // 3) Get the service provider
            var serviceProvider = _host.Services;

            // 4) Show the Main Window
            var mainWindow = serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }
    }
}
