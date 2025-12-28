using FileAutomationSuite.Infrastructure.Interfaces;
using FileAutomationSuite.Infrastructure.Models;
using FileAutomationSuite.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace FileAutomationSuite.Core.Excel
{
    public class ProcessExcel : IExcelBcpService
    {
        //Excel to transefer into DataTable with inferred data types
        #region
        public Dictionary<string, string> ReadExcelColumnsWithDataType(string excelPath)
        {
            var columns = new Dictionary<string, string>();

            using var package = new ExcelPackage(new FileInfo(excelPath));
            var ws = package.Workbook.Worksheets[0];

            int colCount = ws.Dimension.Columns;
            int rowCount = ws.Dimension.Rows;

            for (int col = 1; col <= colCount; col++)
            {
                string columnName = ws.Cells[1, col].Text.Trim();

                if (string.IsNullOrEmpty(columnName))
                    continue;

                columnName = columnName.Replace(" ", "_").Replace("-", "_");

                object sampleValue = null;

                // 🔍 Find first non-empty value (row 2 onwards)
                for (int row = 2; row <= rowCount; row++)
                {
                    if (ws.Cells[row, col].Value != null &&
                        !string.IsNullOrWhiteSpace(ws.Cells[row, col].Text))
                    {
                        sampleValue = ws.Cells[row, col].Value;
                        break;
                    }
                }

                string sqlType = InferSqlDataType(sampleValue);
                columns.Add(columnName, sqlType);
            }

            return columns;
        }
        public void ExcelToBcpTxtSafe(string excelPath, string outputTxtPath)
        {
            ExcelPackage.License.SetNonCommercialOrganization("Basic Use");

            using var package = new ExcelPackage(new FileInfo(excelPath));
            var ws = package.Workbook.Worksheets[0];

            int rows = ws.Dimension.Rows;
            int cols = ws.Dimension.Columns;

            Directory.CreateDirectory(Path.GetDirectoryName(outputTxtPath)!);

            using var writer = new StreamWriter(outputTxtPath, false, Encoding.UTF8);

            for (int row = 2; row <= rows; row++) // skip header
            {
                bool hasData = false;

                for (int col = 1; col <= cols; col++)
                {
                    string value = ConvertCellValueForBCP(ws.Cells[row, col].Value);

                    if (value != "\\N")
                        hasData = true;

                    writer.Write(value);

                    if (col < cols)
                        writer.Write("\t");   // TAB delimiter
                }

                if (hasData)
                    writer.Write("\r\n");    // CRLF for BCP
            }
        }
        public string ConvertCellValueForBCP(object? value)
        {
            if (value == null)
                return "\\N";

            switch (value)
            {
                case double d:
                    return Math.Round(d, 2)
                               .ToString("0.00", CultureInfo.InvariantCulture);

                case decimal m:
                    return Math.Round(m, 2)
                               .ToString("0.00", CultureInfo.InvariantCulture);

                case int or long or short:
                    return value.ToString()!;

                case DateTime dt:
                    return dt.ToString("yyyy-MM-dd HH:mm:ss");

                case bool b:
                    return b ? "1" : "0";

                default:
                    string text = value.ToString()!.Trim();
                    if (string.IsNullOrEmpty(text))
                        return "\\N";

                    return text
                        .Replace("\t", " ")
                        .Replace("\r", " ")
                        .Replace("\n", " ");
            }
        }

        public string InferSqlDataType(object value)
        {
            if (value == null)
                return "VARCHAR(Max)";

            if (decimal.TryParse(value.ToString(), out _))
                return "DECIMAL(18,2)";

            if (int.TryParse(value.ToString(), out _))
                return "INT";

            if (DateTime.TryParse(value.ToString(), out _))
                return "DATETIME";

            if (bool.TryParse(value.ToString(), out _))
                return "BIT";

            return "VARCHAR(Max)";
        }
        public void ConvertExcelToTxt(string excelPath, string txtPath)
        {
            using var package = new ExcelPackage(new FileInfo(excelPath));
            var ws = package.Workbook.Worksheets[0];

            using var writer = new StreamWriter(txtPath, false, Encoding.UTF8);

            for (int row = 2; row <= ws.Dimension.Rows; row++)
            {
                for (int col = 1; col <= ws.Dimension.Columns; col++)
                {
                    writer.Write(ws.Cells[row, col].Text);
                    if (col < ws.Dimension.Columns)
                        writer.Write("\t");
                }
                writer.WriteLine();
            }
        }
        #endregion


        //Excel to Fileter Mutiple Sheet Using column Name 
        #region

        public async Task ProcessExcelAsync(string inputFilePath,string outputFilePath,string filterColumnName,bool includeOriginal,List<int> selectedColumns)
        {
            using var package = new ExcelPackage(new FileInfo(inputFilePath));
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Get filter column index
            int filterColIndex = -1;
            for (int col = 1; col <= colCount; col++)
            {
                if (worksheet.Cells[1, col].Text.Trim()
                    .Equals(filterColumnName, StringComparison.OrdinalIgnoreCase))
                {
                    filterColIndex = col;
                    break;
                }
            }

            if (filterColIndex == -1)
                throw new Exception($"Filter column '{filterColumnName}' not found.");

            // Read grouped data + formatting
            var groupedData = new ConcurrentDictionary<string, List<CellFormatInfo[]>>();

            for (int row = 2; row <= rowCount; row++)
            {
                var cellValue = worksheet.Cells[row, filterColIndex].Text;
                if (string.IsNullOrEmpty(cellValue)) continue;

                var rowFormat = selectedColumns
                    .Select(c => ReadCellFormat(worksheet.Cells[row, c]))
                    .ToArray();

                groupedData.AddOrUpdate(
                    cellValue,
                    new List<CellFormatInfo[]> { rowFormat },
                    (key, list) =>
                    {
                        list.Add(rowFormat);
                        return list;
                    });
            }

            using var outputPackage = new ExcelPackage();

            // -----------------------------------------------------
            // 1. Include original sheet
            // -----------------------------------------------------
            if (includeOriginal)
            {
                var wsOriginal = outputPackage.Workbook.Worksheets.Add("Original");

                // Header
                for (int c = 0; c < selectedColumns.Count; c++)
                {
                    ApplyCellFormat(
                        wsOriginal.Cells[1, c + 1],
                        ReadCellFormat(worksheet.Cells[1, selectedColumns[c]])
                    );

                    wsOriginal.Column(c + 1).Width =
                        worksheet.Column(selectedColumns[c]).Width;
                }

                // Data rows
                int newRow = 2;
                for (int row = 2; row <= rowCount; row++)
                {
                    for (int c = 0; c < selectedColumns.Count; c++)
                    {
                        ApplyCellFormat(
                            wsOriginal.Cells[newRow, c + 1],
                            ReadCellFormat(worksheet.Cells[row, selectedColumns[c]])
                        );
                    }
                    newRow++;

                }

                ExcelHelper.ApplyBordersAndFilters(wsOriginal, selectedColumns.Count, newRow - 1);
            }

            // -----------------------------------------------------
            // 2. Grouped sheets (Alphabetical order)
            // -----------------------------------------------------
            var orderedKeys = groupedData.Keys.OrderBy(k => k).ToList();

            foreach (var key in orderedKeys)
            {
                string sheetName = key.Length > 31 ? key.Substring(0, 31) : key;
                var ws = outputPackage.Workbook.Worksheets.Add(sheetName);

                // Header
                for (int c = 0; c < selectedColumns.Count; c++)
                {
                    ApplyCellFormat(
                        ws.Cells[1, c + 1],
                        ReadCellFormat(worksheet.Cells[1, selectedColumns[c]])
                    );

                    ws.Column(c + 1).Width =
                        worksheet.Column(selectedColumns[c]).Width;
                }

                // Rows
                var rows = groupedData[key];
                int rowIndex = 2;

                foreach (var rowData in rows)
                {
                    for (int c = 0; c < rowData.Length; c++)
                    {
                        ApplyCellFormat(ws.Cells[rowIndex, c + 1], rowData[c]);
                    }
                    rowIndex++;
                }

                ExcelHelper.ApplyBordersAndFilters(ws, selectedColumns.Count, rowIndex - 1);
            }

            // Save file
            var fi = new FileInfo(outputFilePath);
            if (fi.Exists) fi.Delete();

            await outputPackage.SaveAsAsync(fi);
        }



        public void ApplyCellFormat(ExcelRange outputCell, CellFormatInfo format)
        {
            outputCell.Value = format.Value;

            outputCell.Style.HorizontalAlignment = format.HorizontalAlignment;
            outputCell.Style.VerticalAlignment = format.VerticalAlignment;
            outputCell.Style.Numberformat.Format = format.NumberFormat;

            // Font
            outputCell.Style.Font.Bold = format.Bold;
            outputCell.Style.Font.Italic = format.Italic;
            outputCell.Style.Font.Size = format.FontSize;
            outputCell.Style.Font.Name = format.FontName;

            if (format.FontColor != null)
                outputCell.Style.Font.Color.SetColor(format.FontColor.Value);

            // 🔥 IMPORTANT: Apply PatternType before background color
            if (format.BackgroundColor != null)
            {
                outputCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                outputCell.Style.Fill.BackgroundColor.SetColor(format.BackgroundColor.Value);
            }
        }

        public CellFormatInfo ReadCellFormat(ExcelRange cell)
        {
            return new CellFormatInfo
            {
                Value = cell.Value,

                HorizontalAlignment = cell.Style.HorizontalAlignment,
                VerticalAlignment = cell.Style.VerticalAlignment,
                NumberFormat = cell.Style.Numberformat.Format,

                Bold = cell.Style.Font.Bold,
                Italic = cell.Style.Font.Italic,
                FontSize = cell.Style.Font.Size,
                FontName = cell.Style.Font.Name,

                FontColor = GetColor(cell.Style.Font.Color),
                BackgroundColor = GetColor(cell.Style.Fill.BackgroundColor)

            };
        }

        public Color? GetColor(ExcelColor excelColor)
        {
            if (excelColor == null)
                return null;

            // 1. RGB hex (most common)
            if (!string.IsNullOrEmpty(excelColor.Rgb))
            {
                return ColorTranslator.FromHtml("#" + excelColor.Rgb);
            }

            // 2. Indexed color (int, not nullable)
            if (excelColor.Indexed > 0)
            {
                try
                {
                    // Convert index to KnownColor safely
                    return Color.FromKnownColor((KnownColor)excelColor.Indexed);
                }
                catch
                {
                    return null;
                }
            }

            // 3. Theme color (EPPlus does NOT convert automatically)
            if (excelColor.Theme > 0)
            {
                // Cannot resolve without theme lookup → return null safely
                return null;
            }

            return null;
        }

        
        #endregion
    }
}
