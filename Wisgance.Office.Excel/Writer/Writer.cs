using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Wisgance.Office.Excel.General;
using Wisgance.Reflection;

namespace Wisgance.Office.Excel.Writer
{
    public partial class Write
    {
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objects"></param>
        /// <param name="sheetName">generated sheet name, Do not set empty value for this parameters</param>
        /// <param name="headerNames"></param>
        /// <returns></returns>
        public Stream Do<T>(List<T> objects, string sheetName, ExcelHeaderList headerNames)
        {
            var stream = new MemoryStream();

            sheetName = string.IsNullOrEmpty(sheetName) ? "Wisgance_Sheet" : sheetName;

            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetName);

                // Create Styles and Insert into Workbook
                var stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);

                var relId = workbookPart.GetIdOfPart(worksheetPart);

                var workbook = new Workbook();
                var fileVersion = new FileVersion { ApplicationName = "Microsoft Office Excel" };

                var data = objects.Select(o => (object)o).ToList();

                if (headerNames == null)
                {
                    headerNames = new ExcelHeaderList();
                    foreach (var o in ObjUtility.GetPropertyInfo(objects[0]))
                    {
                        headerNames.Add(o, o);
                    }
                }

                var sheetData = CreateSheetData(data, headerNames, stylesPart);

                var worksheet = new Worksheet();

                var numCols = headerNames.Count;
                var width = 20;//headerNames.Max(h => h.Length) + 5;

                var columns = new Columns();
                for (var col = 0; col < numCols; col++)
                {
                    var c = CreateColumnData((UInt32)col + 1, (UInt32)numCols + 1, width);

                    if (c != null) columns.Append(c);
                }
                worksheet.Append(columns);

                var sheets = new Sheets();
                var sheet = new Sheet { Name = sheetName, SheetId = 1, Id = relId };

                sheets.Append(sheet);
                workbook.Append(fileVersion);
                workbook.Append(sheets);

                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                document.WorkbookPart.Workbook = workbook;
                document.WorkbookPart.Workbook.Save();
                document.Close();
            }

            return stream;
        }
    }
}