using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Wisgance.Office.Excel
{
    public static class Utility
    {
        public static UInt32Value CreateCellFormat(Stylesheet styleSheet, UInt32Value fontIndex, UInt32Value fillIndex, UInt32Value numberFormatId)
        {
            CellFormat cellFormat = new CellFormat();

            if (fontIndex != null)
                cellFormat.FontId = fontIndex;

            if (fillIndex != null)
                cellFormat.FillId = fillIndex;

            if (numberFormatId != null)
            {
                cellFormat.NumberFormatId = numberFormatId;
                cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            }

            styleSheet.CellFormats.Append(cellFormat);

            UInt32Value result = styleSheet.CellFormats.Count;
            styleSheet.CellFormats.Count++;
            return result;
        }

        public static UInt32Value CreateFill(Stylesheet styleSheet, System.Drawing.Color fillColor)
        {
            PatternFill patternFill =
                new PatternFill(
                    new ForegroundColor()
                    {
                        Rgb = new HexBinaryValue()
                        {
                            Value =
                            System.Drawing.ColorTranslator.ToHtml(
                                System.Drawing.Color.FromArgb(
                                    fillColor.A,
                                    fillColor.R,
                                    fillColor.G,
                                    fillColor.B)).Replace("#", "")
                        }
                    });

            patternFill.PatternType = fillColor ==
                        System.Drawing.Color.White ? PatternValues.None : PatternValues.LightDown;

            Fill fill = new Fill(patternFill);

            styleSheet.Fills.Append(fill);

            UInt32Value result = styleSheet.Fills.Count;
            styleSheet.Fills.Count++;
            return result;
        }

        public static UInt32Value CreateFont(Stylesheet styleSheet, string fontName, double? fontSize, bool isBold, System.Drawing.Color foreColor)
        {

            Font font = new Font();

            if (!string.IsNullOrEmpty(fontName))
            {
                FontName name = new FontName()
                {
                    Val = fontName
                };
                font.Append(name);
            }

            if (fontSize.HasValue)
            {
                FontSize size = new FontSize()
                {
                    Val = fontSize.Value
                };
                font.Append(size);
            }

            if (isBold == true)
            {
                Bold bold = new Bold();
                font.Append(bold);
            }


            Color color = new Color()
            {
                Rgb = new HexBinaryValue()
                {
                    Value =
                        System.Drawing.ColorTranslator.ToHtml(
                            System.Drawing.Color.FromArgb(
                                foreColor.A,
                                foreColor.R,
                                foreColor.G,
                                foreColor.B)).Replace("#", "")
                }
            };
            font.Append(color);

            styleSheet.Fonts.Append(font);
            UInt32Value result = styleSheet.Fonts.Count;
            styleSheet.Fonts.Count++;
            return result;
        }

        public static Stylesheet CreateStylesheet()
        {
            var ss = new Stylesheet();

            var fts = new Fonts();
            var ftn = new FontName { Val = "Arial" };
            var ftsz = new FontSize { Val = 11 };
            var ft = new DocumentFormat.OpenXml.Spreadsheet.Font { FontName = ftn, FontSize = ftsz };
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;


            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fills.Count = (uint)fills.ChildElements.Count;

            var borders = new Borders();
            var border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            var csfs = new CellStyleFormats();
            var cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 };
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;

            // dd/mm/yyyy is also Excel style index 14

            uint iExcelIndex = 164;
            var nfs = new NumberingFormats();
            var cfs = new CellFormats();

            cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
            cfs.Append(cf);

            var nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "dd/mm/yyyy hh:mm:ss" };
            nfs.Append(nf);

            cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);


            iExcelIndex = 165;
            nfs = new NumberingFormats();
            cfs = new CellFormats();

            cf = new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 };
            cfs.Append(cf);

            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "MMM yyyy" };
            nfs.Append(nf);

            cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);


            iExcelIndex = 170;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.0000" };
            nfs.Append(nf);
            cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            // #,##0.00 is also Excel style index 4
            iExcelIndex = 171;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "#,##0.00" };
            nfs.Append(nf);
            cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            // @ is also Excel style index 49
            iExcelIndex = 172;
            nf = new NumberingFormat { NumberFormatId = iExcelIndex, FormatCode = "@" };
            nfs.Append(nf);
            cf = new CellFormat
            {
                NumberFormatId = nf.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = true
            };
            cfs.Append(cf);

            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            var css = new CellStyles();
            var cs = new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 };
            css.Append(cs);
            css.Count = (uint)css.ChildElements.Count;
            ss.Append(css);

            var dfs = new DifferentialFormats { Count = 0 };
            ss.Append(dfs);

            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium9",
                DefaultPivotStyle = "PivotStyleLight16"
            };
            ss.Append(tss);

            return ss;
        }

        public static string IntToAlpha(int x)
        {
            int lowChar;
            StringBuilder result = new StringBuilder();
            do
            {
                lowChar = (x - 1) % 26;
                x = (x - 1) / 26;
                result.Insert(0, (char)(lowChar + 65));
            } while (x > 0);
            return result.ToString();
        }
    }
}
