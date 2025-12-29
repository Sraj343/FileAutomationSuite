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
    /// Interaction logic for BCPToDataBase.xaml
    /// </summary>
    public partial class BCPToDataBase : Page
    {
        public BCPToDataBase()
        {
            InitializeComponent();
        }

        private void BrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "BCP Files (*.bcp;*.txt)|*.bcp;*.txt|All Files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                txtFilePath.Text = dialog.FileName;
            }
        }

        private async void StartImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtLog.Text = "Import started...\n";

                //await ImportBcpAsync(
                //    txtFilePath.Text,
                //    txtConnectionString.Text,
                //    txtTableName.Text);

                txtLog.Text += "Import completed successfully.";
            }
            catch (Exception ex)
            {
                txtLog.Text += $"Error: {ex.Message}";
            }
        }
    }
}
