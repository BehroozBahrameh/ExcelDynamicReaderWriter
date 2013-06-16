using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wisgance.Office.Excel.General;

namespace Wisgance.Office.Excel.Reader
{
    public partial class Read
    {
        // key      : custom header name
        // value    : property original name
        public static List<T> ReadObjFromExel<T>(Stream stream, ExcelHeaderList pattern, string sheetName) where T : new()
        {
            var result = new List<T>();

            var columnHeaders = GetRowValues(stream, sheetName, "1");

            var hasHeader = columnHeaders.Any(header => pattern != null && pattern.Any(i => i.Key.ToLower().Trim() == header.ToLower().Trim()));

            var excelHeaders = new List<string>();

            for (var i = 1; i <= columnHeaders.Count; i++)
                excelHeaders.Add(Utility.IntToAlpha(i));

            var rowsCount = excelHeaders.Select(header => GetColumnValues(stream, sheetName, header.ToString()).Count).Concat(new[] { 0 }).Max();

            for (var i = hasHeader ? 1 : 0; i < rowsCount; i++)
            {
                var obj = new T();

                #region has header
                if (hasHeader)
                {
                    foreach (var propName in columnHeaders.Where(p => pattern.Any(c => c.Key.ToLower() == p.ToLower())))
                    {
                        var prop = "";

                        var singleOrDefault = pattern.SingleOrDefault(k => k.Key.ToLower() == propName.ToLower());
                        if (singleOrDefault != null)
                            prop = singleOrDefault.Value;

                        var value = GetCellData(stream, sheetName,
                                                   string.Format("{0}{1}", excelHeaders[columnHeaders.IndexOf(propName)],
                                                                 i + 1));

                        if (string.IsNullOrEmpty(prop)) continue;
                        try
                        {
                            var propertyInfo = obj.GetType().GetProperty(prop);
                            propertyInfo.SetValue(obj, value, null);
                        }
                        catch (Exception)
                        { }
                    }
                }
                #endregion

                #region has not hrader
                else
                {
                    var values = GetRowValues(stream, sheetName, (i + 1).ToString());

                    for (var j = 0; j < pattern.Count; j++)
                    {
                        var prop = pattern[j].Value;

                        try
                        {
                            var propertyInfo = obj.GetType().GetProperty(prop);
                            propertyInfo.SetValue(obj, values[j], null);
                        }
                        catch (Exception) { }
                    }
                }
                #endregion

                result.Add(obj);
            }

            return result;
        }

        public static List<T> ReadObjFromExel<T>(Stream stream, ExcelHeaderList pattern) where T : new()
        {
            return ReadObjFromExel<T>(stream, pattern, "");
        }
    }
}
