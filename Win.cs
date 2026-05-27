using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;

namespace ST_ImportExcel
{
    class Program
    {
        static void Main()
        {
            try
            {
                string folderPath = @"C:\DataFiles";

                string oracleConnStr =
                    "Provider=OraOLEDB.Oracle;Data Source=ORCL;User Id=USER;Password=PWD;";

                // Get latest Excel file
                string latestFile = Directory.GetFiles(folderPath, "*.xlsx")
                                             .OrderByDescending(f => File.GetLastWriteTime(f))
                                             .FirstOrDefault();

                if (latestFile == null)
                {
                    Console.WriteLine("No Excel files found.");
                    return;
                }

                string fileName = Path.GetFileName(latestFile);

                using (OleDbConnection oracleConn =
                    new OleDbConnection(oracleConnStr))
                {
                    oracleConn.Open();

                    // Check file already processed
                    string checkSql =
                        "SELECT COUNT(*) FROM FILE_LOG WHERE FILE_NAME = ?";

                    using (OleDbCommand checkCmd =
                        new OleDbCommand(checkSql, oracleConn))
                    {
                        checkCmd.Parameters.AddWithValue(
                            "FILE_NAME", fileName);

                        int count =
                            Convert.ToInt32(checkCmd.ExecuteScalar());

                        if (count > 0)
                        {
                            Console.WriteLine(
                                fileName + " already processed.");

                            return;
                        }
                    }

                    // Excel connection
                    string excelConnStr =
                        @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                        "Data Source=" + latestFile + ";" +
                        "Extended Properties='Excel 12.0 Xml;HDR=YES;'";

                    using (OleDbConnection excelConn =
                        new OleDbConnection(excelConnStr))
                    {
                        excelConn.Open();

                        // Get sheet names
                        DataTable sheets =
                            excelConn.GetOleDbSchemaTable(
                                OleDbSchemaGuid.Tables,
                                null);

                        // FIRST TAB
                        string sheet1 =
                            sheets.Rows[0]["TABLE_NAME"].ToString();

                        // SECOND TAB
                        string sheet2 =
                            sheets.Rows[1]["TABLE_NAME"].ToString();

                        // LOAD FIRST TAB
                        LoadSheetData(
                            excelConn,
                            oracleConn,
                            sheet1,
                            "SALES_DATA",
                            fileName);

                        // LOAD SECOND TAB
                        LoadSheetData(
                            excelConn,
                            oracleConn,
                            sheet2,
                            "CUSTOMER_DATA",
                            fileName);
                    }

                    // Insert log
                    string logSql =
                        "INSERT INTO FILE_LOG " +
                        "(FILE_NAME, LOAD_DATE) " +
                        "VALUES (?, SYSDATE)";

                    using (OleDbCommand logCmd =
                        new OleDbCommand(logSql, oracleConn))
                    {
                        logCmd.Parameters.AddWithValue(
                            "FILE_NAME", fileName);

                        logCmd.ExecuteNonQuery();
                    }

                    Console.WriteLine(
                        fileName + " loaded successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    "Error : " + ex.Message);
            }
        }

        static void LoadSheetData(
            OleDbConnection excelConn,
            OleDbConnection oracleConn,
            string sheetName,
            string tableName,
            string fileName)
        {
            string excelQuery =
                "SELECT * FROM [" + sheetName + "]";

            using (OleDbCommand excelCmd =
                new OleDbCommand(excelQuery, excelConn))
            {
                using (OleDbDataReader reader =
                    excelCmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string id =
                            reader["ID"].ToString();

                        string name =
                            reader["NAME"].ToString();

                        string amount =
                            reader["AMOUNT"].ToString();

                        string insertSql =
                            "INSERT INTO " + tableName +
                            " (FILE_NAME, ID, NAME, AMOUNT) " +
                            "VALUES (?, ?, ?, ?)";

                        using (OleDbCommand insertCmd =
                            new OleDbCommand(insertSql, oracleConn))
                        {
                            insertCmd.Parameters.AddWithValue(
                                "FILE_NAME", fileName);

                            insertCmd.Parameters.AddWithValue(
                                "ID", id);

                            insertCmd.Parameters.AddWithValue(
                                "NAME", name);

                            insertCmd.Parameters.AddWithValue(
                                "AMOUNT", amount);

                            insertCmd.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
    }
}
