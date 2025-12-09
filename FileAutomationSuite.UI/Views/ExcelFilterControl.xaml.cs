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
    /// Interaction logic for ExcelFilterControl.xaml
    /// </summary>
    public partial class ExcelFilterControl : Page
    {
        public ExcelFilterControl()
        {
            InitializeComponent();
        }

        // ================== UPLOAD EXCEL ==================
        private void UploadExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";

            if (dialog.ShowDialog() == true)
            {
                txtInputPath.Text = dialog.FileName;
                LoadExcel(dialog.FileName);
            }
        }

        // Load Excel into DataGrid
        private void LoadExcel(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                DataTable dt = new DataTable();

                // Read header
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    dt.Columns.Add(sheet.Cells[1, col].Text);

                // Read rows
                for (int row = 2; row <= sheet.Dimension.End.Row; row++)
                {
                    DataRow dr = dt.NewRow();
                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                        dr[col - 1] = sheet.Cells[row, col].Text;

                    dt.Rows.Add(dr);
                }

                ExcelGrid.ItemsSource = dt.DefaultView;
            }
        }

        // ================== SELECT OUTPUT PATH ==================
        private void SelectOutputPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Select Folder"
            };

            if (dialog.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(dialog.FileName);
                txtOutputPath.Text = folderPath;
            }
        }


        // ================== SUBMIT ==================
        private void Submit_Click(object sender, RoutedEventArgs e)
        {
            string inputFile = txtInputPath.Text;
            string outputFolder = txtOutputPath.Text;
            string filterColumnName = txtFilterColumn.Text;
            bool isOriginal = chkIsOriginal.IsChecked == true;

            if (string.IsNullOrWhiteSpace(inputFile))
            {
                MessageBox.Show("Please upload an Excel file.");
                return;
            }

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                MessageBox.Show("Please select an output folder.");
                return;
            }

            MessageBox.Show(
                $"File: {inputFile}\n" +
                $"Output: {outputFolder}\n" +
                $"Filter Column: {filterColumnName}\n" +
                $"IsOriginalSheet: {isOriginal}",
                "Submit Clicked");
        }
    }
}
