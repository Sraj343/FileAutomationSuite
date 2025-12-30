using FileAutomationSuite.Infrastructure.Interfaces;
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

namespace FileAutomationSuite.UI.Views
{
    /// <summary>
    /// Interaction logic for ExcelFilterControl.xaml
    /// </summary>
    public partial class ExcelFilterControl : Page
    {
        private readonly IExcelService _excelService;

        // ColumnName -> ColumnIndex mapping
        private Dictionary<string, int> _columnIndexMap = new();

        public ExcelFilterControl()
        {
            InitializeComponent();
            _excelService = (IExcelService)App.ServiceProvider.GetService(typeof(IExcelService))!;
        }

        // ================== UPLOAD EXCEL ==================
        private void UploadExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
            };

            if (dialog.ShowDialog() == true)
            {
                txtInputPath.Text = dialog.FileName;
                LoadExcel(dialog.FileName);
            }
        }

        // ================== LOAD EXCEL & COLUMNS ==================
        private void LoadExcel(string filePath)
        {
            lstColumns.ItemsSource = null;
            _columnIndexMap.Clear();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var sheet = package.Workbook.Worksheets[0];

            if (sheet?.Dimension == null)
            {
                MessageBox.Show("Excel sheet is empty.");
                return;
            }

            List<string> columnNames = new();

            // Read header row
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                string columnName = sheet.Cells[1, col].Text;

                if (!string.IsNullOrWhiteSpace(columnName))
                {
                    columnNames.Add(columnName);
                    _columnIndexMap[columnName] = col; // 1-based index
                }
            }

            lstColumns.ItemsSource = columnNames;
        }

        // ================== SELECT OUTPUT PATH ==================
        private void SelectOutputPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "Select Folder"
            };

            if (dialog.ShowDialog() == true)
            {
                txtOutputPath.Text = Path.GetDirectoryName(dialog.FileName);
            }
        }

        private void ChkSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            lstColumns.SelectAll();
        }


        private void ChkSelectAll_Unchecked(object sender, RoutedEventArgs e)
        {
            lstColumns.UnselectAll();
        }


        private void LstColumns_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstColumns.Items.Count == 0)
                return;

            chkSelectAllColumns.Checked -= ChkSelectAll_Checked;
            chkSelectAllColumns.Unchecked -= ChkSelectAll_Unchecked;

            chkSelectAllColumns.IsChecked =
                lstColumns.SelectedItems.Count == lstColumns.Items.Count;

            chkSelectAllColumns.Checked += ChkSelectAll_Checked;
            chkSelectAllColumns.Unchecked += ChkSelectAll_Unchecked;
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

            // ✅ Get selected column names
            var selectedColumnNames = lstColumns.SelectedItems
                                     .Cast<string>()
                                     .ToList();


            if (!selectedColumnNames.Any())
            {
                MessageBox.Show("Please select at least one column.");
                return;
            }

            // ✅ Convert names → column indexes
            List<int> selectedColumnIndexes = selectedColumnNames
                                              .Select(name => _columnIndexMap[name])
                                              .ToList();

            MessageBox.Show(
                $"File: {inputFile}\n" +
                $"Output: {outputFolder}\n" +
                $"Filter Column: {filterColumnName}\n" +
                $"Selected Columns: {string.Join(", ", selectedColumnNames)}",
                "Submit Clicked");

            // 🔥 Call service
            _excelService.ProcessExcelAsync(
                inputFile,
                outputFolder,
                filterColumnName,
                isOriginal,
                selectedColumnIndexes);
        }
    }
}
