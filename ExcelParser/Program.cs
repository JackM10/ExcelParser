using System.IO;
using static System.Console;
using ExcelDataReader;

namespace _1ListingConsoleApp
{
    class Program
    {
        private static string _pathToXlsx = @"D:\1.xlsx";

        static void Main(string[] args)
        {
            ExcelDataReaderImplementation();
        }

        private static void ExcelDataReaderImplementation()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(_pathToXlsx, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    //do
                    //{
                    //    while (reader.Read())
                    //    {
                    //        WriteLine(reader.GetString(1));
                    //    }   
                    //} while (reader.NextResult());

                    var resaltAsDataSet = reader.AsDataSet();
                    // to take data - resaltAsDataSet.Tables[30].Columns[1].Table.Rows[2].ItemArray // .Table[i] - tab index, .Columns[i] - column index, .Rows[i] - row index
                }
            }
        }
    }
}
