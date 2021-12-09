using System;
using System.Collections.Generic;
using System.IO;
using ExcelGenerator;

namespace ConsoleAppOpenXmlSpreadsheetWithTable
{
    class Program
    {
        static void Main(string[] args)
        {
            var headerRow = new List<string> { "Company", "Country", "Fruit" };
            var exportRowList = new List<ExcelExportRow>();
            exportRowList.Add(new ExcelExportRow() { Company ="Spanish Oranges A", Country = "Spain", Fruit ="Oranges"});
            exportRowList.Add(new ExcelExportRow() { Company = "Mexican Oranges Ltd", Country = "Mexico", Fruit = "Oranges" });
            exportRowList.Add(new ExcelExportRow() { Company = "Medit. Bananas", Country = "Spain", Fruit = "Bananas" });
            exportRowList.Add(new ExcelExportRow() { Company = "Mexican Bananas", Country = "Mexico", Fruit = "Bananas" });
            exportRowList.Add(new ExcelExportRow() { Company = "Mexican Sweet Oranges", Country = "Mexico", Fruit = "Oranges" });
            using (var memoryStream = ExcelFileGenerator.GenerateArticleListWithImagesExcelFile(headerRow, exportRowList, false))
            {
                using (var fileStream = new FileStream("OpenXmlSpreadsheet.xlsx", FileMode.Create))
                {
                    memoryStream.CopyTo(fileStream);
                }
            }
            using (var memoryStream = ExcelFileGenerator.GenerateArticleListWithImagesExcelFile(headerRow, exportRowList, true))
            {
                using (var fileStream = new FileStream("OpenXmlSpreadsheetWithTable.xlsx", FileMode.Create))
                {
                    memoryStream.CopyTo(fileStream);
                }
            }
        }
    }
}
