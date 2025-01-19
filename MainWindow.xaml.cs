using ExcelReport.Services;  // for HollingerReportService
using Microsoft.Extensions.DependencyInjection; // for GetRequiredService
using System.Windows; // for MessageBox

namespace ExcelReport
{
    public partial class MainWindow : Window
    {
        private readonly HollingerReportService _reportService;

        // We'll inject the service from the DI container:
        public MainWindow(HollingerReportService reportService)
        {
            InitializeComponent();
            _reportService = reportService;
        }

        private void ExcelReportButton_Click(object sender, RoutedEventArgs e)
        {
            // We'll just call the service and pass a file path:
            string filePath = @"C:\HollingerReports\HollingerBoxSummery.xlsx";
            _reportService.BuildAndSaveCompleteWorkbook(filePath);
            // _reportService.BuildAndSaveCombinedWorkbook(filePath);
            MessageBox.Show($"Report created at:\n{filePath}");
        }
    }
}
