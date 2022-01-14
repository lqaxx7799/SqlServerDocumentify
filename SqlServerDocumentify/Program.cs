using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace SqlServerDocumentify
{
    class Program
    {
        private const string ConnectionString = "Server=171.244.15.68;Database=Novanet.Ad.V5;User Id=novanetdev;Password=N0vaNetDv1231!@!#;";
        //private const string ConnectionString = "Server=localhost;Database=ecis_db_v2;User Id=sa;Password=1234567;";

        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();
            Console.WriteLine("Starting...");

            var excelpath = Path.Combine(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory), $@"novanet_schema_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
            var fileName = new FileInfo(excelpath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excel = new ExcelPackage(fileName);

            // generate worksheet
            var wsTableDescription = excel.Workbook.Worksheets.Add("Table_Description");
            wsTableDescription.Cells["A1:E1"].Merge = true;
            wsTableDescription.Cells["A1"].Value = "Mô tả nội dung dữ liệu CMS";
            wsTableDescription.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wsTableDescription.Cells["A1"].Style.Font.Bold = true;
            wsTableDescription.Cells["A3"].Value = "STT";
            wsTableDescription.Cells["A3"].Style.Font.Bold = true;
            wsTableDescription.Cells["B3"].Value = "Tên bảng";
            wsTableDescription.Cells["B3"].Style.Font.Bold = true;
            wsTableDescription.Cells["C3"].Value = "Mô tả";
            wsTableDescription.Cells["C3"].Style.Font.Bold = true;
            wsTableDescription.Cells["D3"].Value = "Hiện có dữ liệu (Y/N)";
            wsTableDescription.Cells["D3"].Style.Font.Bold = true;

            var wsTableColumnDescription = excel.Workbook.Worksheets.Add("Table_Column_Description");
            wsTableColumnDescription.Cells["A1:E1"].Merge = true;
            wsTableColumnDescription.Cells["A1"].Value = "Mô tả các trường theo từng bảng";
            wsTableColumnDescription.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wsTableColumnDescription.Cells["A1"].Style.Font.Bold = true;
            wsTableColumnDescription.Cells["A3"].Value = "STT";
            wsTableColumnDescription.Cells["A3"].Style.Font.Bold = true;
            wsTableColumnDescription.Cells["B3"].Value = "Tên bảng";
            wsTableColumnDescription.Cells["B3"].Style.Font.Bold = true;
            wsTableColumnDescription.Cells["C3"].Value = "Tên trường";
            wsTableColumnDescription.Cells["C3"].Style.Font.Bold = true;
            wsTableColumnDescription.Cells["D3"].Value = "Kiểu dữ liệu";
            wsTableColumnDescription.Cells["D3"].Style.Font.Bold = true;
            wsTableColumnDescription.Cells["E3"].Value = "Mô tả";
            wsTableColumnDescription.Cells["E3"].Style.Font.Bold = true;

            var wsColumnDescription = excel.Workbook.Worksheets.Add("Column_Description");
            wsColumnDescription.Cells["A1:E1"].Merge = true;
            wsColumnDescription.Cells["A1"].Value = "Mô tả các trường CMS";
            wsColumnDescription.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            wsColumnDescription.Cells["A1"].Style.Font.Bold = true;
            wsColumnDescription.Cells["A3"].Value = "STT";
            wsColumnDescription.Cells["A3"].Style.Font.Bold = true;
            wsColumnDescription.Cells["B3"].Value = "Tên trường";
            wsColumnDescription.Cells["B3"].Style.Font.Bold = true;
            wsColumnDescription.Cells["C3"].Value = "Kiểu dữ liệu";
            wsColumnDescription.Cells["C3"].Style.Font.Bold = true;
            wsColumnDescription.Cells["D3"].Value = "Ý nghĩa";
            wsColumnDescription.Cells["D3"].Style.Font.Bold = true;

            using var connection = new SqlConnection(ConnectionString);
            connection.Open();
            Console.WriteLine("Connection opened!");

            var tables = GetTables(connection);
            Console.WriteLine($"Found {tables.Count} tables");

            var currentTableRow = 4;
            var currentTableColumnRow = 4;

            var allColumns = new List<TableSchema>();

            for (var i = 0; i < tables.Count; i++)
            {
                if (i % 5 == 0)
                {
                    Console.WriteLine($"Read schemas of {i + 1}/{tables.Count}");
                }

                var table = tables[i];
                wsTableDescription.Cells[$"A{currentTableRow}"].Value = currentTableRow - 3;
                wsTableDescription.Cells[$"B{currentTableRow}"].Value = table.TableName;

                var schema = GetSchema(connection, table.TableName);
                allColumns.AddRange(schema);

                wsTableColumnDescription.Cells[$"A{currentTableColumnRow}:E{currentTableColumnRow}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                wsTableColumnDescription.Cells[$"A{currentTableColumnRow}:E{currentTableColumnRow}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                wsTableColumnDescription.Cells[$"A{currentTableColumnRow}"].Value = currentTableRow - 3;
                wsTableColumnDescription.Cells[$"B{currentTableColumnRow}"].Value = table.TableName;

                currentTableRow += 1;

                for (var j = 0; j < schema.Count; j++)
                {
                    var column = schema[j];
                    wsTableColumnDescription.Cells[$"C{currentTableColumnRow}"].Value = column.ColumnName;
                    wsTableColumnDescription.Cells[$"D{currentTableColumnRow}"].Value = column.DataType;
                    currentTableColumnRow += 1;
                }

                // prevent overloading database
                Thread.Sleep(100);
            }

            allColumns = allColumns.OrderBy(x => x.ColumnName).ToList();
            var currentColumnRow = 4;
            for (var j = 0; j < allColumns.Count; j++)
            {
                var column = allColumns[j];
                wsColumnDescription.Cells[$"A{currentColumnRow}"].Value = currentColumnRow - 3;
                wsColumnDescription.Cells[$"B{currentColumnRow}"].Value = column.ColumnName;
                wsColumnDescription.Cells[$"C{currentColumnRow}"].Value = column.DataType;
                currentColumnRow += 1;
            }

            wsTableDescription.Column(2).AutoFit();
            wsTableDescription.Column(3).Width = 40;
            wsTableDescription.Column(4).AutoFit();

            wsTableColumnDescription.Column(2).AutoFit();
            wsTableColumnDescription.Column(3).AutoFit();
            wsTableColumnDescription.Column(4).AutoFit();
            wsTableColumnDescription.Column(5).Width = 40;

            wsColumnDescription.Column(2).AutoFit();
            wsColumnDescription.Column(3).AutoFit();
            wsColumnDescription.Column(4).Width = 40;

            Stream stream = File.Create(excelpath);
            excel.SaveAs(stream);
            stream.Close();

            stopwatch.Stop();
            Console.WriteLine($"Finised after {stopwatch.ElapsedMilliseconds}ms.");

            Console.WriteLine("Press any key to close!");
            Console.ReadKey();
        }

        static List<TableInformation> GetTables(SqlConnection connection)
        {
            var tablesQuery = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME";
            var command = new SqlCommand(tablesQuery, connection);
            var reader = command.ExecuteReader();
            var result = new List<TableInformation>();
            try
            {
                while (reader.Read())
                {
                    var row = new TableInformation
                    {
                        TableCatalog = (string)reader["TABLE_CATALOG"],
                        TableSchema = (string)reader["TABLE_SCHEMA"],
                        TableName = (string)reader["TABLE_NAME"],
                        TableType = (string)reader["TABLE_TYPE"]
                    };
                    result.Add(row);
                }
                return result;
            }
            finally
            {
                // Always call Close when done reading.
                reader.Close();
            }
        }

        static List<TableSchema> GetSchema(SqlConnection connection, string tableName)
        {
            var tablesQuery = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME=@tableName ORDER BY COLUMN_NAME";
            var command = new SqlCommand(tablesQuery, connection);
            command.Parameters.AddWithValue("@tableName", tableName);

            var reader = command.ExecuteReader();
            var result = new List<TableSchema>();
            try
            {
                while (reader.Read())
                {
                    var row = new TableSchema
                    {
                        TableName = (string)reader["TABLE_NAME"],
                        ColumnName = (string)reader["COLUMN_NAME"],
                        DataType = (string)reader["DATA_TYPE"],
                        CharacterMaximumLength = reader["CHARACTER_MAXIMUM_LENGTH"] != DBNull.Value ? (int)reader["CHARACTER_MAXIMUM_LENGTH"] : null
                    };
                    result.Add(row);
                }
                return result;
            }
            finally
            {
                // Always call Close when done reading.
                reader.Close();
            }
        }
    }
}
