using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Wisgance.Office.Excel.Writer.CellTypes
{
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            DataType = CellValues.Date;
            CellReference = header + index;
            StyleIndex = 1;
            CellValue = new CellValue { Text = dateTime.ToOADate().ToString() };
        }
    }
}
