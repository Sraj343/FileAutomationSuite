using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileAutomationSuite.Core.SQL
{
    class SQLProcess
    {

        static void CreateTableFromExcel(string connStr, string tableName, Dictionary<string, string> columns)
        {
            using var conn = new SqlConnection(connStr);
            conn.Open();

            var sql = new StringBuilder();

            sql.AppendLine($"IF OBJECT_ID('{tableName}','U') IS NULL");
            sql.AppendLine($"BEGIN");
            sql.AppendLine($"CREATE TABLE [{tableName}] (");

            int index = 0;
            foreach (var column in columns)
            {
                sql.Append($"    [{column.Key}] {column.Value}");

                if (index < columns.Count - 1)
                    sql.Append(",");

                sql.AppendLine();
                index++;
            }

            sql.AppendLine(")");
            sql.AppendLine("END");

            Console.WriteLine("Generated SQL:");
            Console.WriteLine(sql.ToString());

            using var cmd = new SqlCommand(sql.ToString(), conn);
            cmd.ExecuteNonQuery();
        }

        static void CreateTableInTargetFromSource(string sourceConn, string targetConn, string tableName)
        {
            string createScript = GenerateCreateTableScript(sourceConn, tableName);

            using var conn = new SqlConnection(targetConn);
            conn.Open();

            new SqlCommand(createScript, conn).ExecuteNonQuery();
        }

        static string GenerateCreateTableScript(string connStr, string tableName)
        {
            using var conn = new SqlConnection(connStr);
            conn.Open();

            string sql = $@"
DECLARE @sql NVARCHAR(MAX)='CREATE TABLE [{tableName}] (';
SELECT @sql=@sql+CHAR(13)+'['+c.name+'] '+
       TYPE_NAME(c.user_type_id)+
       CASE WHEN c.max_length=-1 THEN '(MAX)'
            WHEN c.max_length>0 THEN '('+CAST(c.max_length AS VARCHAR)+')'
            ELSE '' END+
       CASE WHEN c.is_nullable=1 THEN ' NULL,' ELSE ' NOT NULL,' END
FROM sys.columns c WHERE c.object_id=OBJECT_ID('{tableName}');
SELECT LEFT(@sql,LEN(@sql)-1)+');';";

            return new SqlCommand(sql, conn).ExecuteScalar().ToString();
        }
    }
}
