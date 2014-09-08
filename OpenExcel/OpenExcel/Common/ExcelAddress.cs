namespace OpenExcel.Common
{
    using System;
    using System.Text.RegularExpressions;

    public static class ExcelAddress
    {
        private static Regex _rgxAddress = new Regex(@"(\$)?([A-Z]+)(\$)?([0-9]+)(:((\$)?([A-Z]+)(\$)?([0-9]+)))?$", RegexOptions.Compiled);

        public static string ColumnIndexToName(uint col)
        {
            col--;
            int num = 1;
            uint num2 = 0x1a;
            uint num3 = 0x19;
            uint num4 = 0;
            while (col > num3)
            {
                num2 *= 0x1a;
                num4 = num3 + 1;
                num3 += num2;
                num++;
            }
            col -= num4;
            char[] chArray = new char[num];
            for (int i = 0; i < num; i++)
            {
                chArray[(num - i) - 1] = (char) (0x41 + (col % 0x1a));
                col /= 0x1a;
            }
            return new string(chArray);
        }

        public static uint ColumnNameToIndex(string colName)
        {
            if (string.IsNullOrEmpty(colName))
            {
                throw new ArgumentException("Invalid columnName [" + colName + "]");
            }
            int length = colName.Length;
            uint num2 = 0x1a;
            uint num3 = 0x19;
            uint num4 = 0;
            for (int i = 0; i < length; i++)
            {
                num2 *= 0x1a;
                if (i < (length - 1))
                {
                    num4 = num3 + 1;
                }
                num3 += num2;
            }
            uint num6 = 0;
            for (int j = 0; j < length; j++)
            {
                num6 *= 0x1a;
                num6 +=(uint)( colName[j] - 'A');
            }
            num6 += num4;
            return (num6 + 1);
        }

        public static uint GetColumn(string address)
        {
            Match match = _rgxAddress.Match(address);
            if (!match.Success)
            {
                throw new ArgumentException("Invalid Excel address");
            }
            return ColumnNameToIndex(match.Groups[2].Value);
        }

        public static RowColumn ToRowColumn(string address)
        {
            Match match = _rgxAddress.Match(address);
            if (!match.Success)
            {
                throw new ArgumentException("Invalid Excel address");
            }
            uint num = uint.Parse(match.Groups[4].Value);
            if (num == 0)
            {
                throw new ArgumentException("Row cannot be zero");
            }
            uint num2 = ColumnNameToIndex(match.Groups[2].Value);
            return new RowColumn { Row = num, Column = num2 };
        }
    }
}

