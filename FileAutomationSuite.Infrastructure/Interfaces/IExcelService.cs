using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Infrastructure.Interfaces
{
    public interface IExcelService
    {
        Dictionary<string, string> ReadExcelColumnsWithDataType(string excelPath);
        void ExcelToBcpTxtSafe(string excelPath, string outputTxtPath);

        Task ProcessExcelAsync(string inputFilePath, string outputFilePath, string filterColumnName, bool includeOriginal, List<int> selectedColumns);
    }
}
