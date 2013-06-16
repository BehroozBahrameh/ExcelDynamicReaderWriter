using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Wisgance.Office.Excel.General;
using Wisgance.Office.Excel.Writer.CellTypes;
using Wisgance.Reflection;

namespace Wisgance.Office.Excel.Writer
{
    public partial class Write
    {
        //private SheetData CreateSheetData<T>(List<T> objects, ExcelHeaderList headerTitles, WorkbookStylesPart stylesPart)
        //{
        //    SheetData sheetData = new SheetData();

        //    if (objects != null)
        //    {
        //        List<string> fields = ObjUtility.GetPropertyInfo<T>();

        //        var az = new List<Char>(Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => (Char)i).ToArray());

        //        List<Char> headers = az.GetRange(0, fields.Count);

        //        Row header = new Row();

        //        int index = 1;

        //        header.RowIndex = (uint)index;

        //        foreach (ExcelHeader keyValuePair in headerTitles.Where(i => i.HeaderType == ExcelHeaderType.SingleProperty))
        //        {
        //            HeaderCell c = new HeaderCell(headers[headerTitles.ToList().IndexOf(keyValuePair)].ToString(),
        //                                            keyValuePair.Value,
        //                                            index,
        //                                            stylesPart.Stylesheet,
        //                                            System.Drawing.Color.DodgerBlue,
        //                                            12,
        //                                            true);

        //            header.Append(c);
        //        }

        //        sheetData.Append(header);

        //        for (int i = 0; i < objects.Count; i++)
        //        {
        //            index++;

        //            var selectedObj = objects[i];

        //            var r = new Row { RowIndex = (uint)index };

        //            int headerIndex = 0;

        //            foreach (ExcelHeader keyValuePair in headerTitles.OrderBy(h => (byte)h.HeaderType))
        //            {
        //                PropertyInfo myf = selectedObj.GetType().GetProperty(keyValuePair.Key);

        //                if (myf != null)
        //                {
        //                    object obj = myf.GetValue(selectedObj, null);
        //                    if (obj != null)
        //                    {
        //                        if (obj.GetType() == typeof(string))
        //                        {
        //                            TextCell c = new TextCell(headers[headerIndex].ToString(),
        //                            obj.ToString(),
        //                            index);

        //                            r.Append(c);
        //                        }
        //                        else if (obj.GetType() == typeof(bool))
        //                        {
        //                            string value = (bool)obj ? "Yes" : "No";
        //                            TextCell c = new TextCell(headers[headerIndex].ToString(),
        //                            value,
        //                            index);

        //                            r.Append(c);
        //                        }
        //                        else if (obj.GetType() == typeof(DateTime))
        //                        {
        //                            string value = ((DateTime)obj).ToOADate().ToString();
        //                            DateCell c = new DateCell(headers[headerIndex].ToString(),
        //                            (DateTime)obj,
        //                            index);

        //                            r.Append(c);
        //                        }
        //                        else if (obj.GetType() == typeof(decimal) || obj.GetType() == typeof(double))
        //                        {
        //                            FormatedNumberCell c = new FormatedNumberCell(headers[headerIndex].ToString(),
        //                            obj.ToString(),
        //                            index);

        //                            r.Append(c);
        //                        }
        //                        else if (obj.GetType() == typeof(Dictionary<string, string>) && keyValuePair.HeaderType == ExcelHeaderType.ListProperty)
        //                        {
        //                            Dictionary<string, string> objList = (obj as Dictionary<string, string>);

        //                            foreach (var VARIABLE in objList)
        //                            {
        //                                TextCell c = new TextCell(
        //                                    headers[headerIndex].ToString(),
        //                                    VARIABLE.Value,
        //                                    index);
        //                                r.Append(c);

        //                                HeaderCell hc = new HeaderCell(
        //                                    headers[headerIndex].ToString(),
        //                                    VARIABLE.Key,
        //                                    1,
        //                                    stylesPart.Stylesheet,
        //                                    System.Drawing.Color.Brown,
        //                                    12,
        //                                    true);

        //                                header.Append(hc);

        //                                headerIndex++;
        //                            }

        //                        }
        //                        else
        //                        {
        //                            long value;
        //                            if (long.TryParse(obj.ToString(), out value))
        //                            {
        //                                NumberCell c = new NumberCell(headers[headerIndex].ToString(),
        //                                obj.ToString(),
        //                                index);

        //                                r.Append(c);
        //                            }
        //                            else
        //                            {
        //                                TextCell c = new TextCell(headers[headerIndex].ToString(),
        //                                obj.ToString(),
        //                                index);

        //                                r.Append(c);
        //                            }
        //                        }
        //                    }
        //                }
        //                headerIndex++;
        //            }
        //            sheetData.Append(r);
        //        }

        //        index++;
        //    }

        //    return sheetData;
        //}

        private SheetData CreateSheetData(IReadOnlyList<object> objects, ExcelHeaderList headerTitles, WorkbookStylesPart stylesPart)
        {
            var sheetData = new SheetData();

            if (objects != null)
            {
                var header = new Row();
                var index = 1;

                header.RowIndex = (uint)index;

                foreach (ExcelHeader keyValuePair in headerTitles.Where(i => i.HeaderType == ExcelHeaderType.SingleProperty))
                {
                    HeaderCell c = new HeaderCell(Utility.IntToAlpha(headerTitles.ToList().IndexOf(keyValuePair) + 1),
                                                    keyValuePair.Value,
                                                    index,
                                                    stylesPart.Stylesheet,
                                                    System.Drawing.Color.DodgerBlue,
                                                    12,
                                                    true);

                    header.Append(c);
                }

                sheetData.Append(header);

                for (int i = 0; i < objects.Count; i++)
                {
                    index++;

                    var selectedObj = objects[i];

                    var row = new Row { RowIndex = (uint)index };

                    int headerIndex = 1;

                    foreach (ExcelHeader keyValuePair in headerTitles.OrderBy(h => (byte)h.HeaderType))
                    {
                        PropertyInfo myf = selectedObj.GetType().GetProperty(keyValuePair.Key);

                        if (myf != null)
                        {
                            object obj = myf.GetValue(selectedObj, null);
                            if (obj != null)
                            {
                                CreateDataCell(row, obj, headerIndex, ref index);
                                headerIndex++;
                            }
                        }
                    }
                    sheetData.Append(row);
                }

                index++;
            }

            return sheetData;
        }

        private Column CreateColumnData(UInt32 startColumnIndex, UInt32 endColumnIndex, double columnWidth)
        {
            Column column;
            column = new Column();
            column.Min = startColumnIndex;
            column.Max = endColumnIndex;
            column.Width = columnWidth;
            column.CustomWidth = true;
            return column;
        }

        private void CreateDataCell(Row row, object obj, int headerIndex, ref int index)
        {
            string header = Utility.IntToAlpha(headerIndex);

            if (obj is string)
            {
                row.Append(new TextCell(header, obj.ToString(), index));
            }
            else if (obj is bool)
            {
                string value = (bool)obj ? "Yes" : "No";
                row.Append(new TextCell(header, value, index));
            }
            else if (obj is DateTime)
            {
                string value = ((DateTime)obj).ToOADate().ToString();
                row.Append(new DateCell(header, (DateTime)obj, index));
            }
            else if (obj is decimal || obj is double)
            {
                row.Append(new FormatedNumberCell(header, obj.ToString(), index));
            }
            else if (obj.GetType().GetInterface("ICollection", true) != null)
            {
                var collection = obj as System.Collections.ICollection;
                if (collection != null)
                    foreach (var item in collection)
                    {
                        CreateDataCell(row, item, headerIndex, ref index);
                        headerIndex++;
                    }
            }
            else
            {
                long value;
                if (long.TryParse(obj.ToString(), out value))
                {
                    row.Append(new NumberCell(header, obj.ToString(), index));
                }
                else
                {
                    row.Append(new TextCell(header, obj.ToString(), index));
                }
            }
        }
    }
}