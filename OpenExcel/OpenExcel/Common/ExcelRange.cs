namespace OpenExcel.Common
{
    using OpenExcel.Common.RangeParser;
    using System;
    using System.Runtime.InteropServices;

    public static class ExcelRange
    {
        public static RangeComponents Parse(string range)
        {
            return new RangeComponents(range);
        }

        public static string Translate(string range, int rowDelta, int colDelta)
        {
            if (range == null)
            {
                return null;
            }
            return TranslateInternal(Parse(range), 0, 0, rowDelta, colDelta, true);
        }

        public static string TranslateForSheetChange(string range, SheetChange sheetChange, string currentSheetName)
        {
            if (range == null)
            {
                return null;
            }
            RangeComponents er = Parse(range);
            if (((er.SheetName != "") || (currentSheetName != sheetChange.SheetName)) && !(er.SheetName == sheetChange.SheetName))
            {
                return range;
            }
            return TranslateInternal(er, sheetChange.RowStart, sheetChange.ColumnStart, sheetChange.RowDelta, sheetChange.ColumnDelta, false);
        }

        private static string TranslateInternal(RangeComponents er, uint rowStart, uint colStart, int rowDelta, int colDelta, bool followAbsoluteRefs)
        {
            string str = null;
            string str2 = null;
            bool exceededBounds = false;
            bool flag2 = false;
            if (er.Cell1Error != "")
            {
                str = er.Cell1Error;
            }
            else
            {
                RowColumn column = er.Cell1RowColumn;
                str = TranslateInternal(er.Cell1RowDollar, column.Row, er.Cell1ColDollar, column.Column, rowStart, colStart, rowDelta, colDelta, followAbsoluteRefs, out exceededBounds);
            }
            if (er.Cell2 != "")
            {
                if (er.Cell2Error != "")
                {
                    str2 = er.Cell2Error;
                }
                else
                {
                    RowColumn column2 = er.Cell2RowColumn;
                    str2 = TranslateInternal(er.Cell2RowDollar, column2.Row, er.Cell2ColDollar, column2.Column, rowStart, colStart, rowDelta, colDelta, followAbsoluteRefs, out flag2);
                }
            }
            string str3 = "";
            if (er.SheetName != "")
            {
                str3 = str3 + er.EscapedSheetName + "!";
            }
            if (exceededBounds && (flag2 || (str2 == null)))
            {
                return (str3 + "#REF!");
            }
            str3 = str3 + str;
            if (str2 != null)
            {
                str3 = str3 + ":" + str2;
            }
            return str3;
        }

        private static string TranslateInternal(string rowDollar, uint rowIdx, string colDollar, uint colIdx, uint rowStart, uint colStart, int rowDelta, int colDelta, bool followAbsoluteRefs, out bool exceededBounds)
        {
            int num = (int) rowIdx;
            int num2 = (int) colIdx;
            if ((num >= rowStart) && ((rowDollar != "$") || !followAbsoluteRefs))
            {
                num += rowDelta;
            }
            if ((num2 >= colStart) && ((colDollar != "$") || !followAbsoluteRefs))
            {
                num2 += colDelta;
            }
            exceededBounds = false;
            if (((num < 1) || (num > ExcelConstraints.MaxRows)) || ((num2 < 1) || (num2 > ExcelConstraints.MaxColumns)))
            {
                exceededBounds = true;
            }
            num = Math.Max(1, Math.Min(num, ExcelConstraints.MaxRows));
            num2 = Math.Max(1, Math.Min(num2, ExcelConstraints.MaxColumns));
            return string.Concat(new object[] { rowDollar, ExcelAddress.ColumnIndexToName((uint) num2), colDollar, num });
        }
    }
}

