using FileAutomationSuite.Infrastructure.Interfaces;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace FileAutomationSuite.UI.Views
{
    /// <summary>
    /// Interaction logic for ExcelToBCP.xaml
    /// </summary>
    public partial class ExcelToBCP : Page
    {

        private readonly IExcelService _excelService;
        public ExcelToBCP()
        {
            InitializeComponent();
            _excelService = (IExcelService)App.ServiceProvider.GetService(typeof(IExcelService))!;
        }

        private void BrowseExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
                TxtExcelPath.Text = dialog.FileName;
        }

        private void BrowseTxt_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Text Files (*.txt)|*.txt"
            };

            if (dialog.ShowDialog() == true)
                TxtOutputPath.Text = dialog.FileName;
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(TxtExcelPath.Text))
            {
                MessageBox.Show("Please select a valid Excel file.");
                return;
            }

            if (string.IsNullOrWhiteSpace(TxtOutputPath.Text))
            {
                MessageBox.Show("Please select output TXT path.");
                return;
            }

            try
            {
                _excelService.ExcelToBcpTxtSafe(TxtExcelPath.Text, TxtOutputPath.Text);
                TxtLog.Text = "✔ Excel successfully converted to BCP-safe TXT file.";
            }
            catch (Exception ex)
            {
                TxtLog.Text = ex.Message;
            }
        }

        private void ReadColumns_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(TxtExcelPath.Text))
            {
                MessageBox.Show("Please select an Excel file.");
                return;
            }

            var columns = _excelService.ReadExcelColumnsWithDataType(TxtExcelPath.Text);

            var sb = new StringBuilder();
            sb.AppendLine("Detected Columns & SQL Data Types");
            sb.AppendLine("--------------------------------");

            foreach (var col in columns)
                sb.AppendLine($"{col.Key} : {col.Value}");

            TxtLog.Text = sb.ToString();
        }
    }
}
