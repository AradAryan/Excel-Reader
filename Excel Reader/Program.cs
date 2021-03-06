﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Reader
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileName = @"";
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
        }
    }
}
