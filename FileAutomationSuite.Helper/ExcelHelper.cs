using FileAutomationSuite.Infrastructure.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Utility
{
    public static class ExcelHelper
    {
        public static void SaveFileToMergedPath(string inputFile, string directoryPath, string backupFilePath)
        {
            // 1. Merge directory and backup folder paths
            string finalFolderPath = Path.Combine(directoryPath, backupFilePath);

            // 2. Create directory if it doesn't exist
            if (!Directory.Exists(finalFolderPath))
                Directory.CreateDirectory(finalFolderPath);

            // 3. Extract file name
            string fileName = Path.GetFileName(inputFile);

            // 4. Create final file save path
            string destinationPath = Path.Combine(finalFolderPath, fileName);

            // 5. Copy file to merged folder
            File.Copy(inputFile, destinationPath, true);
        }

        public static List<List<ExcelCellData>> ReadExcelWithStyles(string filePath)
        {
            
            var result = new List<List<ExcelCellData>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = package.Workbook.Worksheets.First();
                int rows = ws.Dimension.End.Row;
                int cols = ws.Dimension.End.Column;

                // Read headers
                var headers = new List<string>();
                for (int c = 1; c <= cols; c++)
                {
                    headers.Add(ws.Cells[1, c].Text);
                }

                // Read row-wise data
                for (int r = 2; r <= rows; r++)
                {
                    var rowList = new List<ExcelCellData>();

                    for (int c = 1; c <= cols; c++)
                    {
                        var cell = ws.Cells[r, c];

                        rowList.Add(new ExcelCellData()
                        {
                            Header = headers[c - 1],
                            Value = cell.Value,
                            NumberFormat = cell.Style.Numberformat.Format,
                            FontName = cell.Style.Font.Name,
                            FontSize = cell.Style.Font.Size,
                            FontColor = ColorTranslator.ToHtml(cell.Style.Font.Color.Rgb != null
                                ? Color.FromArgb(Convert.ToInt32(cell.Style.Font.Color.Rgb, 16))
                                : Color.Black),
                            BackgroundColor = cell.Style.Fill.BackgroundColor.Rgb != null
                                ? "#" + cell.Style.Fill.BackgroundColor.Rgb
                                : "#FFFFFF"
                        });
                    }

                    result.Add(rowList);
                }
            }

            return result;
        }

        public static List<ExcelHeaderInfo> ReadExcelHeadersWithStyles(string filePath)
        {
            var result = new List<ExcelHeaderInfo>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];
                if (sheet == null) return result;

                int headerRow = -1;

                // 🔍 Step 1: Detect header row (first non-empty row)
                for (int row = 1; row <= sheet.Dimension.End.Row; row++)
                {
                    bool anyCellHasValue = false;
                    for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                    {
                        if (!string.IsNullOrWhiteSpace(sheet.Cells[row, col].Text))
                        {
                            anyCellHasValue = true;
                            break;
                        }
                    }
                    if (anyCellHasValue)
                    {
                        headerRow = row;
                        break;
                    }
                }

                if (headerRow == -1) return result;

                // 🔍 Step 2: Read headers with formatting
                for (int col = 1; col <= sheet.Dimension.End.Column; col++)
                {
                    var cell = sheet.Cells[headerRow, col];

                    if (string.IsNullOrWhiteSpace(cell.Text))
                        continue;

                    var font = cell.Style.Font;
                    var fill = cell.Style.Fill;

                    result.Add(new ExcelHeaderInfo
                    {
                        HeaderName = cell.Text,
                        Value = cell.Text,
                        FontColor = font.Color?.Rgb,
                        BackgroundColor = fill.BackgroundColor?.Rgb,
                        FontSize = font.Size,
                        Bold = font.Bold,
                        Italic = font.Italic
                    });
                }
            }

            return result;
        }

        public static string BuildConnectionString(string server, string db, string user, string pass)
        {
            return $"Server={server};Database={db};User Id={user};Password={pass};Encrypt=True;TrustServerCertificate=True;";
        }

        public static IConfiguration GetConfig(string configFilePath)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Path.GetDirectoryName(configFilePath))
                .AddJsonFile(Path.GetFileName(configFilePath), optional: false, reloadOnChange: true);

            return builder.Build();
        }

        public static async Task ApplyBordersAndFilters(ExcelWorksheet ws, int colCount, int rowCount)
        {
            // Borders
            using (var range = ws.Cells[1, 1, rowCount, colCount])
            {
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            // AutoFilter
            ws.Cells[1, 1, rowCount, colCount].AutoFilter = true;

            // Autofit rows
            ws.Cells.AutoFitColumns();
        }
    }
}
