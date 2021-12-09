using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGenerator
{
    public class ExcelExportRow
    {
        public string Company { get; set; }
        public string Country { get; set; }
        public string Fruit { get; set; }
    }

    public static class ExcelFileGenerator
    {
        private const int normalFontStyleIndex = 0;
        private const int boldFontStyleIndex = 1;

        public static MemoryStream GenerateArticleListWithImagesExcelFile(List<string> headerRow, List<ExcelExportRow> excelExportRoListw, bool insertTable)
        {
            MemoryStream memoryStream = new MemoryStream();

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                Workbook workbook = new Workbook();
                workbookPart.Workbook = workbook;

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");

                Stylesheet stylesheet = StyleSheet();
                workbookStylesPart.Stylesheet = stylesheet;
                workbookStylesPart.Stylesheet.Save();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                Worksheet worksheet = new Worksheet();
                SheetData sheetData = new SheetData();

                Sheets sheets = new Sheets();

                //get the string name of the columns
                string[] excelColumnNamesTitle = new string[5];
                for (int n = 0; n < 5; n++)
                    excelColumnNamesTitle[n] = GetExcelColumnName(n);

                var columns = new Columns();
                columns.Append(new Column() { Min = 1, Max = 1, CustomWidth = true, Width = 24 });
                columns.Append(new Column() { Min = 2, Max = 3, CustomWidth = true, Width = 12 });
                worksheet.Append(columns);

                if (excelExportRoListw is not null)
                {
                    Row rowTitle = new Row() { RowIndex = (UInt32Value)1 };
                    for (int c = 0; c < headerRow.Count; c++)
                    {
                        AppendTextCell(excelColumnNamesTitle[c], headerRow[c], rowTitle, boldFontStyleIndex);
                    }
                    sheetData.Append(rowTitle);

                    uint rowNumber = 2;
                    foreach (var excelExportRow in excelExportRoListw)
                    {
                        var spreadsheetRow = new Row() { RowIndex = (UInt32Value)rowNumber };
                        AppendTextCell(excelColumnNamesTitle[0], excelExportRow.Company, spreadsheetRow, normalFontStyleIndex);
                        AppendTextCell(excelColumnNamesTitle[2], excelExportRow.Country, spreadsheetRow, normalFontStyleIndex);
                        AppendTextCell(excelColumnNamesTitle[3], excelExportRow.Fruit, spreadsheetRow, normalFontStyleIndex);
                        sheetData.Append(spreadsheetRow);
                        rowNumber++;
                    }

                    worksheet.Append(sheetData);
                    worksheetPart.Worksheet = worksheet;
                }

                // Table
                if (insertTable)
                    DefineTable(worksheetPart, 1, 1 + excelExportRoListw.Count, 0, 2);

                // Wrap up and save
                worksheetPart.Worksheet.Save();

                Sheet sheet = new Sheet() { Name = "Export", SheetId = (UInt32Value)1, Id = workbookPart.GetIdOfPart(worksheetPart) };
                sheets.Append(sheet);
                workbook.Append(sheets);
                workbook.Save();
                document.Close();
            }

            memoryStream.Position = 0;
            return memoryStream;
        }


        private static void DefineTable(WorksheetPart worksheetPart, int rowMin, int rowMax, int colMin, int colMax)
        {
            #region Table
            string rangeReference = $"{GetExcelColumnName(colMin)}{rowMin}:{GetExcelColumnName(colMax)}{rowMax}";
            int tableNo = 1;

            Table table = new Table()
            {
                Id = (UInt32)tableNo,
                Name = "Table" + tableNo,
                DisplayName = "Table" + tableNo,
                Reference = rangeReference,
                TotalsRowShown = false
            };

            AutoFilter autoFilter = new AutoFilter() { Reference = rangeReference };

            TableColumns tableColumns = new TableColumns() { Count = (UInt32)(colMax - colMin + 1) };
            for (int i = 0; i < (colMax - colMin + 1); i++)
            {
                tableColumns.Append(new TableColumn() { Id = (UInt32)(i + 1), Name = "Column" + i });
            }

            TableStyleInfo tableStyleInfo = new TableStyleInfo()
            {
                Name = "TableStyleMedium1",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            };

            table.Append(autoFilter);
            table.Append(tableColumns);
            table.Append(tableStyleInfo);
            #endregion

            string tableRefId = "rIdTable1";

            // Add table to worksheet - two different things implictly linked by using the same refId
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>(tableRefId);
            tableDefinitionPart.Table = table;
            
            TableParts tableParts = new TableParts();
            TablePart tablePart = new TablePart() { Id = tableRefId };
            tableParts.Append(tablePart);
            tableParts.Count = 1;
            worksheetPart.Worksheet.Append(tableParts);
        }

        #region Excel Helper funcs
        private static void AppendTextCell(string cellReference, string cellStringValue, Row excelRow, UInt32Value styleIndex)
        {
            //  Add a new Excel Cell to our Row
            Cell cell = new Cell() { CellReference = cellReference, DataType = CellValues.String };
            CellValue cellValue = new CellValue();
            cellValue.Text = cellStringValue;
            cell.StyleIndex = styleIndex;
            cell.Append(cellValue);
            excelRow.Append(cell);
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            if (columnIndex < 26)
                return ((char)('A' + columnIndex)).ToString();

            char firstChar = (char)('A' + (columnIndex / 26) - 1);
            char secondChar = (char)('A' + (columnIndex % 26));

            return string.Format("{0}{1}", firstChar, secondChar);
        }
        #endregion Helper funcs

        #region Stylesheet
        private static Stylesheet StyleSheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            Fonts fonts = new Fonts();
            // 0 - normal fonts
            Font myFont = new Font()
            {
                FontSize = new FontSize() { Val = 11 },
                Color = new Color() { Rgb = HexBinaryValue.FromString("FF000000") },
                FontName = new FontName() { Val = "Arial" }
            };
            fonts.Append(myFont);

            //1 - font bold
            myFont = new Font()
            {
                Bold = new Bold(),
                FontSize = new FontSize() { Val = 11 },
                Color = new Color() { Rgb = HexBinaryValue.FromString("FF000000") },
                FontName = new FontName() { Val = "Arial" }
            };
            fonts.Append(myFont);


            Fills fills = new Fills();
            //default fill
            Fill fill = new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            };
            fills.Append(fill);

            Borders borders = new Borders();
            //normal borders
            Border border = new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);

            CellFormats cellFormats = new CellFormats();
            //0- normalFontStyleIndex for normal cells
            CellFormat cellFormat = new CellFormat()
            {
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                ApplyFill = false
            };
            cellFormats.Append(cellFormat);

            //1 - boldFontStyleIndex for header row
            cellFormat = new CellFormat()
            {
                FontId = 1,
                FillId = 0,
                BorderId = 0,
                ApplyFill = false
            };
            cellFormats.Append(cellFormat);

            styleSheet.Append(fonts);
            styleSheet.Append(fills);
            styleSheet.Append(borders);
            styleSheet.Append(cellFormats);

            TableStyles tableStyles = new TableStyles()
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium1",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            styleSheet.Append(tableStyles);

            return styleSheet;
        }
        #endregion Stylesheet
    }
}
