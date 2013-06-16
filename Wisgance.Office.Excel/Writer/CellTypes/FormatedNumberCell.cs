namespace Wisgance.Office.Excel.Writer.CellTypes
{
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, string text, int index)
            : base(header, text, index)
        {
            StyleIndex = 2;
        }
    }
}
