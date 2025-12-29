using System;
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Dialogs;
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
    /// Interaction logic for DBTableToBCP.xaml
    /// </summary>
    public partial class DBTableToBCP : Page
    {
        public DBTableToBCP()
        {
            InitializeComponent();
        }

        // Browse for folder
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Select Output Folder for BCP file"
            };

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                OutputPathTextBox.Text = dialog.FileName;
            }
        }

        // Clear all fields
        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            ConnectionStringTextBox.Text = string.Empty;
            DatabaseNameTextBox.Text = string.Empty;
            TableNameTextBox.Text = string.Empty;
            OutputPathTextBox.Text = string.Empty;
            FileNameTextBox.Text = string.Empty;
        }

        // Export to BCP (placeholder)
        private void ExportToBcp_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = ConnectionStringTextBox.Text;
            string databaseName = DatabaseNameTextBox.Text;
            string tableName = TableNameTextBox.Text;
            string outputPath = OutputPathTextBox.Text;
            string fileName = FileNameTextBox.Text;

            if (string.IsNullOrWhiteSpace(connectionString) ||
                string.IsNullOrWhiteSpace(databaseName) ||
                string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(outputPath) ||
                string.IsNullOrWhiteSpace(fileName))
            {
                System.Windows.MessageBox.Show("Please fill all fields.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // TODO: Implement your BCP export logic here
            System.Windows.MessageBox.Show($"BCP export started for table {tableName}!", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }

    }
}
