using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wisgance.Reflection;

namespace Wisgance.Office.Excel.Reader
{
    public partial class Read
    {
        public static dynamic ReadObjFromExel(Stream stream)
        {
            var result = new List<object>();
            var columnHeaders = GetRowValues(stream, "", "1");

            var objProp = (columnHeaders.Where(header => !string.IsNullOrEmpty(header))
                .Select(header => new FieldMask()
                        {
                            FieldName=header,
                            FieldType = typeof(string)
                        })).ToList();
            var excelHeaders = new List<string>();

            for (var i = 1; i <= columnHeaders.Count; i++)
                excelHeaders.Add(Utility.IntToAlpha(i));

            //var rowsCount = excelHeaders.Select(header => GetColumnValues(stream, "", header.ToString()).Count).Concat(new[] {0}).Max();

            //for (var i = 1; i < rowsCount; i++)
            long k = 1;
            while (true)
            {
                var obj = MyTypeBuilder.CreateNewObject(objProp);

                var values = GetRowValues(stream, "", (k + 1).ToString());
                if (!values.Any())
                    break;

                for (var c = 0; c < objProp.Count; c++)
                {
                    try
                    {
                        var propertyInfo = obj.GetType().GetProperty(objProp[c].FieldName);
                        propertyInfo.SetValue(obj, values[c], null);

                    }
                    catch (Exception) { }
                }

                result.Add(obj);
                k++;
            }

            return result;
        }
    }
}