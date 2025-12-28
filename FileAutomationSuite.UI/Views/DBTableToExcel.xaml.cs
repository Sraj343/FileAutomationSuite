using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
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
    /// Interaction logic for DBTableToExcel.xaml
    /// </summary>
    public partial class DBTableToExcel : Page
    {
        // Properties for simplicity; in MVVM these would be in a ViewModel
        public string SelectedDatabase { get; set; }
        public string TableName { get; set; }
        public string OutputPath { get; set; }

        // Example database list
        public string[] DatabaseList { get; set; } = { "Database1", "Database2", "Database3" };

        public DBTableToExcel()
        {
            InitializeComponent();
        }

        // Browse button click
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                OutputPath = dialog.FileName;
                // Update binding manually
                OutputPathTextBox.Text = OutputPath;
            }
        }

        // Export to Excel button click
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SelectedDatabase) || string.IsNullOrWhiteSpace(TableName))
            {
                MessageBox.Show("Please select a database and enter a table name.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(OutputPath))
            {
                MessageBox.Show("Please select an output path for the Excel file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                string connectionString = $"Server=YOUR_SERVER_NAME;Database={SelectedDatabase};Trusted_Connection=True;";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand($"SELECT * FROM {TableName}", conn);
                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add(TableName);
                        ws.Cells["A1"].LoadFromDataTable(dt, true);
                        File.WriteAllBytes(OutputPath, package.GetAsByteArray());
                    }
                }

                MessageBox.Show("Export completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting table: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Clear button click
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            SelectedDatabase = null;
            TableName = string.Empty;
            OutputPath = string.Empty;

            DatabaseComboBox.SelectedItem = null;
            TableNameTextBox.Text = string.Empty;
            OutputPathTextBox.Text = string.Empty;
        }
    }
}
