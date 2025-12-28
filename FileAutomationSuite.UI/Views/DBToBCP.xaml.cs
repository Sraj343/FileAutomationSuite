using System;
using System.Collections.Generic;
using System.Diagnostics;
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
    /// Interaction logic for DBToBCP.xaml
    /// </summary>
    public partial class DBToBCP : Page
    {
        public DBToBCP()
        {
            InitializeComponent();
        }


        // Browse button click
        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Output Folder";
                dialog.ShowNewFolderButton = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    OutputFolderTextBox.Text = dialog.SelectedPath;
                }
            }
        }

        // Export button click
        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = ConnectionStringTextBox.Text.Trim();
            string tableName = TableNameTextBox.Text.Trim();
            string outputFolder = OutputFolderTextBox.Text.Trim();
            string bcpFileName = BcpFileNameTextBox.Text.Trim();

            if (string.IsNullOrEmpty(connectionString) ||
                string.IsNullOrEmpty(tableName) ||
                string.IsNullOrEmpty(outputFolder) ||
                string.IsNullOrEmpty(bcpFileName))
            {
                System.Windows.MessageBox.Show("Please fill all fields.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string fullPath = Path.Combine(outputFolder, bcpFileName);

            try
            {
                // Example BCP command
                string bcpCommand = $"bcp {tableName} out \"{fullPath}\" -S . -T -c";

                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", "/c " + bcpCommand)
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    process.WaitForExit();
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    if (!string.IsNullOrEmpty(error))
                        System.Windows.MessageBox.Show(error, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    else
                        System.Windows.MessageBox.Show("BCP export completed successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
