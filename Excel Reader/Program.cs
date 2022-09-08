using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Reader
{
    public static class Program
    {
        static void Main(string[] args)
        {
            const string fileName = @"C:\Users\javad\Desktop\accounts.xlsx";
            var result = Excel.ReadFile(fileName);
            result.Count();
            #region comment
            /*
            const string fileName = @"C:\Users\javad\Desktop\accounts.xlsx";
            StreamReader streamReader = new StreamReader(fileName);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[1];
                for (int RowIndex = 2; RowIndex <= worksheet.Dimension.End.Row; RowIndex++)
                {
                    try
                    {
                        var temp = worksheet.Cells[RowIndex, 1].Value.ToString();
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            */
            #endregion
        }
    }
}
