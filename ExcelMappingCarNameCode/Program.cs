using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMappingCarNameCode
{
    class Program
    {
        static void Main(string[] args)
        {
            String path = ConfigurationManager.AppSettings["FileExcel"].ToString();
            Console.WriteLine("Please Wait...");

            var response = FileToTable(path);

            SqlConnection sqlCon = null;
            String SqlconString = ConfigurationManager.AppSettings["ConnectionString"].ToString();

            using (sqlCon = new SqlConnection(SqlconString))
            {
                sqlCon.Open();
                SqlCommand sql_cmnd = new SqlCommand("SpVMIUpdateCarNameMapping", sqlCon);
                sql_cmnd.CommandType = CommandType.StoredProcedure;
                sql_cmnd.Parameters.AddWithValue("@Xml", SqlDbType.Xml).Value = response.Replace("&", "&amp;");
                sql_cmnd.ExecuteNonQuery();
                sqlCon.Close();
            }
            Console.WriteLine("Insert Table Completed...");
        }
        private static string FileToTable(String pathFile)
        {
            var sb = new System.Text.StringBuilder();
            FileStream stream = File.Open(pathFile, FileMode.Open, FileAccess.Read);
            var i = 0;
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    if (i != 0 && !(reader[0] is null) && !string.IsNullOrWhiteSpace(reader[0].ToString()))
                    {
                        sb.AppendLine("<Group>");

                        sb.Append("<carname_code>");
                        sb.Append(reader[0] is null ? "" : reader[0].ToString());
                        sb.Append("</carname_code>");

                        sb.Append("<carname_desc>");
                        sb.Append(reader[1] is null ? "" : reader[1].ToString());
                        sb.Append("</carname_desc>");

                        sb.Append("<carname_model>");
                        sb.Append(reader[2] is null ? "" : reader[2].ToString());
                        sb.Append("</carname_model>");

                        sb.Append("<carname_cc>");
                        sb.Append(reader[3] is null ? "" : reader[3].ToString());
                        sb.Append("</carname_cc>");

                        sb.Append("<car_option>");
                        sb.Append(reader[4] is null ? "" : reader[4].ToString());
                        sb.AppendLine("</car_option>");

                        sb.AppendLine("</Group>");
                    }
                    i++;
                }
            }

            return "<Carname>" + sb.ToString() + "</Carname>";
        }
    }
}
