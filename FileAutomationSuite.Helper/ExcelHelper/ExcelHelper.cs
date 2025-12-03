using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Helper.ExcelHelper
{
    class ExcelHelper
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

    }
}
