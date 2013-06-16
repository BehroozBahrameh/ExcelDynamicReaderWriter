using DocumentFormat.OpenXml.Spreadsheet;

namespace Wisgance.Office.Excel.Writer.CellTypes
{
    public class FomulaCell : Cell
    {
        public FomulaCell(string header, string text, int index)
        {
            CellFormula = new CellFormula { CalculateCell = true, Text = text };
            DataType = CellValues.Number;
            CellReference = header + index;
            StyleIndex = 2;
        }
    }
}
