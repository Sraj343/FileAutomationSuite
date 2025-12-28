using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
    /// Interaction logic for ExcelToDBTable.xaml
    /// </summary>
    public partial class ExcelToDBTable : Page
    {
        public ExcelToDBTable()
        {
            InitializeComponent();
        }

        // Browse Excel File
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
            dlg.Title = "Select Excel File";

            if (dlg.ShowDialog() == true)
            {
                _excelFilePath = dlg.FileName;
                ExcelFileTextBox.Text = _excelFilePath; // TextBox name from XAML
                AppendLog($"Selected Excel file: {_excelFilePath}");
            }
        }

        // Preview Data (placeholder)
        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_excelFilePath))
            {
                MessageBox.Show("Please select an Excel file first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            AppendLog("Previewing Excel data...");

            // TODO: Load Excel data into a DataGrid
            // Example: Use EPPlus, ClosedXML, or Interop to read Excel
        }

        // Upload to DB (placeholder)
        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_excelFilePath))
            {
                MessageBox.Show("Please select an Excel file first.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string connectionString = DbConnectionTextBox.Text.Trim();
            string tableName = TableNameTextBox.Text.Trim();

            if (string.IsNullOrEmpty(connectionString) || string.IsNullOrEmpty(tableName))
            {
                MessageBox.Show("Please enter database connection string and table name.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            AppendLog($"Uploading data from {_excelFilePath} to table {tableName}...");

            // TODO: Implement actual Excel -> DB logic
        }

        // Append message to status log
        private void AppendLog(string message)
        {
            StatusTextBox.AppendText($"{DateTime.Now:HH:mm:ss} - {message}\n");
            StatusTextBox.ScrollToEnd();
        }
    }
}
