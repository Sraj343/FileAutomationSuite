using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Core.BCP
{
    class BCPProcess
    {

        public void ExportTableUsingBCP(string server, string database, string table, string filePath, string user, string pass)
        {
            string cmd =
                $"bcp \"{database}.dbo.{table}\" out \"{filePath}\" -c -t\"|\" -r\"\\n\" " +
                $"-S {server} -U {user} -P {pass}";

            RunCommand(cmd);
        }

        public void ImportTableUsingBCP(string server, string database, string table, string filePath, string user, string pass)
        {
            string cmd =
                $"bcp \"{database}.dbo.{table}\" in \"{filePath}\" -c -t\"\\t\" -r\"\\n\" " +
                $"-S {server} -U {user} -P {pass}";

            RunCommand(cmd);
        }

        public void RunCommand(string command)
        {
            Console.WriteLine("Running BCP Command:");
            Console.WriteLine(command);

            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/c " + command,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);

            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            Console.WriteLine("BCP OUTPUT:");
            Console.WriteLine(output);

            if (!string.IsNullOrWhiteSpace(error))
            {
                Console.WriteLine("❌ BCP ERROR:");
                Console.WriteLine(error);
            }
        }
    }
}
