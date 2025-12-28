using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Core.BCP
{
    public class BcpExporter
    {

        /// <summary>
        /// Exports an entire SQL table into a BCP text file.
        /// </summary>
        public static string ExportTableToBcp(
            string connectionString,
            string databaseName,
            string tableName,
            string outputDirectory)
        {
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            string filePath = Path.Combine(
                outputDirectory,
                $"{tableName}_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

            string arguments =
                $"{databaseName}.dbo.{tableName} out \"{filePath}\" " +
                "-c -t\"|\" " +           // -c = character mode, -t = | delimiter
                "-S . " +                 // Server name (use '.' for local)
                BuildCredentialArgs(connectionString);

            var psi = new ProcessStartInfo
            {
                FileName = "bcp",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var process = new Process { StartInfo = psi };

            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (!string.IsNullOrEmpty(error))
                throw new Exception("BCP Error: " + error);

            Console.WriteLine(output);

            return filePath;
        }

        private static string BuildCredentialArgs(string connectionString)
        {
            var builder = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(connectionString);

            if (builder.IntegratedSecurity)
                return "-T"; // trusted connection

            return $"-U {builder.UserID} -P {builder.Password}";
        }


        public static bool ImportBcpFile(
        string connectionString,
        string databaseName,
        string tableName,
        string bcpFilePath)
        {
            // BCP command
            string arguments =
                $"{databaseName}.dbo.{tableName} in \"{bcpFilePath}\" " +
                "-c -t\"|\" " +             // -c = character mode, -t = column delimiter
                "-S . " +                   // Server (use '.' for local)
                $"-E " +                    // Use trusted connection (remove if using SQL user)
                BuildCredentialArgs(connectionString);

            var processStartInfo = new ProcessStartInfo
            {
                FileName = "bcp",
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var process = new Process
            {
                StartInfo = processStartInfo
            };

            process.Start();

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (!string.IsNullOrEmpty(error))
                throw new Exception("BCP Error: " + error);

            Console.WriteLine(output);
            return true;
        }
    }
}
