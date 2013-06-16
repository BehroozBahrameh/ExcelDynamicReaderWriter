using System.Collections.Generic;

namespace Wisgance.Office.Excel.General
{
    /// <summary>
    /// key      : custom header name
    /// value    : property original name
    /// </summary>
    public class ExcelHeader
    {
        public string Key { get; set; }

        public string Value { get; set; }

        public ExcelHeaderType HeaderType { get; set; }
    }

    public class ExcelHeaderList : List<ExcelHeader>
    {
        public void Add(string key, string value, ExcelHeaderType type)
        {
            this.Add(new ExcelHeader()
                         {
                             Key = key,
                             Value = value,
                             HeaderType = type
                         });
        }

        public void Add(string key, string value)
        {
            this.Add(new ExcelHeader()
            {
                Key = key,
                Value = value,
                HeaderType = ExcelHeaderType.SingleProperty
            });
        }
    }

    public enum ExcelHeaderType : byte
    {
        SingleProperty = 1,
        ListProperty = 2
    }
}
