using System;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Web.Mvc;
using System.IO;

namespace Excel_Reader
{
    public class Excel : Controller
    {
        private static int startRow;
        private static int endRow;
        private static int startColumn;
        private static int endColumn;

        #region Reader
        private static int fromRow;
        private static int toRow;
        private static int fromColumn;
        private static int toColumn;

        public static List<List<string>> ReadFile(string FilePath, int ToRow, int ToColumn, int WorkSheet = 1, int FromColumn = 0, int FromRow = 0)
        {
            StreamReader streamReader = new StreamReader(FilePath);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[WorkSheet];

                fromRow = FromRow;
                toRow = ToRow;
                fromColumn = FromColumn;
                toColumn = ToColumn;

                return Reader(worksheet);
            }
        }
        public static List<List<string>> ReadFile(string FilePath, int ToColumn, int WorkSheet = 1, int FromColumn = 0, int FromRow = 0)
        {
            StreamReader streamReader = new StreamReader(FilePath);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[WorkSheet];

                fromRow = FromRow;
                fromColumn = FromColumn;
                toColumn = ToColumn;
                toColumn = worksheet.Dimension.End.Column;

                return Reader(worksheet);
            }
        }
        public static List<List<string>> ReadFile(string FilePath, short WorkSheet = 1, int FromColumn = 0, int FromRow = 0)
        {
            StreamReader streamReader = new StreamReader(FilePath);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[WorkSheet];

                fromRow = FromRow;
                fromColumn = FromColumn;
                toColumn = worksheet.Dimension.End.Column;
                toRow = worksheet.Dimension.End.Row;

                return Reader(worksheet);
            }
        }
        public static List<List<string>> ReadFile(string FilePath, int ToRow, int WorkSheet = 1, int FromRow = 0)
        {
            StreamReader streamReader = new StreamReader(FilePath);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[WorkSheet];

                fromRow = FromRow;
                toRow = ToRow;
                fromColumn = 0;
                toColumn = worksheet.Dimension.End.Column;

                return Reader(worksheet);
            }
        }
        public static List<List<string>> ReadFile(string FilePath)
        {
            StreamReader streamReader = new StreamReader(FilePath);
            using (var excel = new ExcelPackage(streamReader.BaseStream))
            {
                var workbook = excel.Workbook;
                var worksheet = excel.Workbook.Worksheets[1];

                fromRow = 0;
                toRow = toColumn = worksheet.Dimension.End.Row;
                fromColumn = 0;
                toColumn = toColumn = worksheet.Dimension.End.Column;

                return Reader(worksheet);
            }
        }
        private static List<List<string>> Reader(ExcelWorksheet worksheet)
        {
            List<List<string>> data = new List<List<string>>();
            for (int RowIndex = fromRow; RowIndex <= toRow; RowIndex++)
            {
                List<string> column = new List<string>
                    {
                        RowIndex.ToString()
                    };
                for (int ColumnIndex = fromColumn; ColumnIndex <= toColumn; ColumnIndex++)
                {
                    try
                    {
                        string value = worksheet.Cells[RowIndex, ColumnIndex].Value.ToString();
                        column.Add(value);
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
                data.Add(column);
            }
            return data;
        }
        #endregion

        public static byte[] WriteFile(string TemplateFilePath, int EndRow, int EndColumn, int WorkSheet = 1, int StartColumn = 0, int StartRow = 0)
        {

            return null;
        }
        public static byte[] WriteFile(string TemplateFilePath, int EndColumn, int WorkSheet = 1, int StartColumn = 0, int StartRow = 0)
        {
            return null;
        }
        public static byte[] WriteFile(string TemplateFilePath, short WorkSheet = 1, int StartColumn = 0, int StartRow = 0)
        {
            return null;
        }
        public static byte[] WriteFile(string TemplateFilePath, int EndRow, int WorkSheet = 1, int StartRow = 0)
        {
            return null;
        }
        public static byte[] WriteFile(string TemplateFilePath, string Path, string FileName, int worksheet = 1)
        {
            using (var fileStream = new FileStream(Path, FileMode.Open, FileAccess.Read))
            {
                using (var excel = new ExcelPackage(fileStream))
                {
                    ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets[worksheet];
                    var test = Excel.ExcelWrite(excelWorksheet, "test");
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }
            return null;
        }

        public static FileContentResult ExcelWrite(ExcelWorksheet worksheet, string FileName)
        {
            List<string> list = new List<string>();

            var RowIndex = 7;
            if (list.Count() > 0)
            {
                foreach (var item in list)
                {
                    var temp = item.Split(',');
                    worksheet.Cells[RowIndex, 3].Value = RowIndex - 6;
                    worksheet.Cells[RowIndex, 6].Value = temp[0];
                    RowIndex++;
                }

                //worksheet.Protection.AllowSort = true;
                //worksheet.Protection.AllowAutoFilter = true;
                //worksheet.Protection.SetPassword("password");
                //worksheet.Protection.IsProtected = true;

                /* Style
                var select = worksheet.SelectedRange[3, 1, --counter, 14];
                select.StyleName = "Style 1";
                select.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                select.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                select = worksheet.SelectedRange[3, 10, counter, 10];
                select.StyleName = "Style 2";
                select.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                select.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                */
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                worksheet.Protection.IsProtected = true;
                worksheet.View.FreezePanes(2, 1); // freeze header row
                worksheet.Protection.AllowSort = true;
                worksheet.Cells[worksheet.Dimension.Address].AutoFilter = true;
                worksheet.Protection.AllowAutoFilter = true;
            }
            return File(excel.GetAsByteArray(), "vnd.openxmlformats-officedocument.spreadsheetml.sheet", FileName);
            //File.WriteAllBytes(@"C:\Users\javad\Desktop\test.xlsx", excel.GetAsByteArray("password"));

        }
    }
}
