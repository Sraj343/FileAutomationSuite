using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Infrastructure.Interfaces
{
    public interface IBCPService
    {
        void ImportTableUsingBCP(string server, string database, string table, string filePath, string user, string pass);
        void ExportTableUsingBCP(string server, string database, string table, string filePath, string user, string pass);
    }
}
