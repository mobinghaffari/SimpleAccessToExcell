using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string accessFilePath = @"C:\Users\www80\Downloads\Compressed\GI_TimeBilling_v4\2.mdb";
        string excelFilePath = @"C:\Users\www80\Downloads\output.xlsx";
        string connectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={accessFilePath};";

        //string connectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={accessFilePath};Persist Security Info=False;";

        try
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable schemaTable = connection.GetSchema("Tables");

                using (var workbook = new XLWorkbook())
                {
                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string tableName = row["TABLE_NAME"].ToString();

                        if (tableName.StartsWith("MSys")) // Skip system tables
                            continue;

                        using (OleDbCommand command = new OleDbCommand($"SELECT * FROM [{tableName}]", connection))
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            var worksheet = workbook.Worksheets.Add(dataTable, tableName);
                        }
                    }

                    workbook.SaveAs(excelFilePath);
                }
            }

            Console.WriteLine($"Data has been successfully exported to {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}