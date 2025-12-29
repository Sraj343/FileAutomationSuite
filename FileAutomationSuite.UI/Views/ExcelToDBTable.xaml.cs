using Microsoft.Win32;
using System;
using System.Windows;
using System.Windows.Controls;

namespace FileAutomationSuite.UI.Views
{
    public partial class ExcelToDBTable : Page
    {
        private string _excelFilePath;

        public ExcelToDBTable()
        {
            InitializeComponent();
        }

        // Browse Excel File
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                Title = "Select Excel File"
            };

            if (dlg.ShowDialog() == true)
            {
                _excelFilePath = dlg.FileName;
                ExcelFileTextBox.Text = _excelFilePath;
                AppendLog($"Selected Excel file: {_excelFilePath}");
            }
        }

        // Preview Data
        private void PreviewButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_excelFilePath))
            {
                MessageBox.Show("Please select an Excel file first.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            AppendLog("Previewing Excel data...");
            // TODO: Load Excel into DataGrid (EPPlus / ClosedXML)
        }

        // Upload to DB
        private void UploadButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_excelFilePath))
            {
                MessageBox.Show("Please select an Excel file first.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            string connectionString = DbConnectionTextBox.Text.Trim();
            string tableName = TableNameTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(connectionString) ||
                string.IsNullOrWhiteSpace(tableName))
            {
                MessageBox.Show("Please enter database connection string and table name.",
                                "Warning",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                return;
            }

            AppendLog($"Uploading data from {_excelFilePath} to table {tableName}...");
            // TODO: Implement Excel → DB insert logic
        }

        // Status log helper
        private void AppendLog(string message)
        {
            StatusTextBox.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
            StatusTextBox.ScrollToEnd();
        }
    }
}
