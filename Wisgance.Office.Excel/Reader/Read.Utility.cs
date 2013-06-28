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
        /// <summary>
        /// Get all specific column's cell data
        /// </summary>
        /// <param name="file">Excel file</param>
        /// <param name="sheetName">selected sheet, if set empty or incorrect sheet name, automatically get first sheet</param>
        /// <param name="reference">column name</param>
        /// <returns></returns>
        private static List<string> GetColumnValues(Stream file, string sheetName, string reference)
        {
            var result = new List<string>();

            //Read Excel File by OpenXml Library
            using (var document = SpreadsheetDocument.Open(file, false))
            {
                var workbook = document.WorkbookPart;

                var theSheet = workbook.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ??
                                 workbook.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);

                var worksheet = (WorksheetPart)(workbook.GetPartById(theSheet.Id));

                var cells = worksheet.Worksheet.Descendants<Cell>().Where(c => GetCellCol(c.CellReference).ToUpper() == reference);

                //get cell data by calling ExtractCellValue function
                result.AddRange(from theCell in cells where theCell != null select ExtractCellValue(theCell, workbook));
            }

            return result;
        }

        /// <summary>
        /// Get all specific row's cell data
        /// </summary>
        /// <param name="file">Excel file</param>
        /// <param name="sheetName">selected sheet, if set empty or incorrect sheet name, automatically get first sheet</param>
        /// <param name="reference">cell name></param>
        /// <returns></returns>
        private static List<string> GetRowValues(Stream file, string sheetName, string reference)
        {
            var result = new List<string>();

            //Read Excel File by OpenXml Library
            using (var document = SpreadsheetDocument.Open(file, false))
            {
                var workbook = document.WorkbookPart;

                //If sheet name dose not exsits, get the first shhet of excel file.
                var theSheet = workbook.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ??
                               workbook.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);

                if (theSheet == null)
                {
                    throw new ArgumentException("NOT EXISTS SHEET!!");
                }
                
                var workSheet = (WorksheetPart)(workbook.GetPartById(theSheet.Id));

                //read all cells in selected row of sheet by passed "reference" argument
                var cells =
                    workSheet.Worksheet.Descendants<Cell>().Where(c => GetCellRow(c.CellReference).ToUpper() == reference);

                //get cell data by calling ExtractCellValue function
                result.AddRange(from theCell in cells where theCell != null select ExtractCellValue(theCell, workbook));
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

        /// <summary>
        /// Extract cell's data by cell's data type
        /// </summary>
        /// <param name="theCell">the cell that we want extract data</param>
        /// <param name="workbook">workbook of selected cell</param>
        /// <returns>cell's data in string format</returns>
        private static string ExtractCellValue(Cell theCell, WorkbookPart workbook)
        {
            string value = theCell.InnerText;

            if (theCell.DataType != null)
            {
                switch (theCell.DataType.Value)
                {
                    case CellValues.SharedString:

                        var stringTable = workbook.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

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
