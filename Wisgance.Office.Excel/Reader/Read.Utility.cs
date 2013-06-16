using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Wisgance.Office.Excel.Reader
{
    public partial class Read
    {
        private static List<string> GetColumnValues(Stream file, string sheetName, string reference)
        {
            var result = new List<string>();

            using (var document = SpreadsheetDocument.Open(file, false))
            {
                var wbPart = document.WorkbookPart;

                var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ??
                                 wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);

                var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                var cells = wsPart.Worksheet.Descendants<Cell>().Where(c => GetCellCol(c.CellReference).ToUpper() == reference);

                result.AddRange(from theCell in cells where theCell != null select ExtractCellValue(theCell, wbPart));
            }

            return result;
        }

        private static List<string> GetRowValues(Stream file, string sheetName, string reference)
        {
            var result = new List<string>();

            using (var document = SpreadsheetDocument.Open(file, false))
            {
                var wbPart = document.WorkbookPart;

                var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                if (theSheet == null)
                {
                    theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);
                }

                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                var cells =
                    wsPart.Worksheet.Descendants<Cell>().Where(c => GetCellRow(c.CellReference).ToUpper() == reference);

                result.AddRange(from theCell in cells where theCell != null select ExtractCellValue(theCell, wbPart));
            }

            return result;
        }

        private static string GetCellData(Stream file, string sheetName, string reference)
        {
            var result = string.Empty;

            using (var document = SpreadsheetDocument.Open(file, false))
            {
                var wbPart = document.WorkbookPart;

                var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ??
                                 wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);

                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                var theCell = wsPart.Worksheet.Descendants<Cell>().SingleOrDefault(c => c.CellReference == reference);

                if (theCell != null)
                    result = ExtractCellValue(theCell, wbPart);
            }
            return result;
        }

        private static string ExtractCellValue(Cell theCell, WorkbookPart wbPart)
        {
            string value = theCell.InnerText;

            if (theCell.DataType != null)
            {
                switch (theCell.DataType.Value)
                {
                    case CellValues.SharedString:

                        var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                        if (stringTable != null)
                        {
                            value = stringTable.SharedStringTable.
                                ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            return value;
        }

        private static string GetCellCol(string reference)
        {
            var index = reference.IndexOfAny(new char[]
                                                 {
                                                     '0',
                                                     '1',
                                                     '2',
                                                     '3',
                                                     '4',
                                                     '5',
                                                     '6',
                                                     '7',
                                                     '8',
                                                     '9'
                                                 }
                );

            return index < 0 ? reference : reference.Substring(0, index);
        }

        private static string GetCellRow(string reference)
        {
            return reference.Replace(GetCellCol(reference), "");
        }
    }
}
