namespace OpenExcel.Common.RangeParser
{
    using OpenExcel.Common;
    using System;
    using System.Text.RegularExpressions;

    public class RangeComponents
    {
        private Match _m;
        private static Regex _rgxRef = new Regex(@"^((?<SheetName>.+)!)?(?<C1>((?<C1ColDollar>\$)?(?<C1Col>[A-Z]+)(?<C1RowDollar>\$)?(?<C1Row>[0-9]+)|(?<C1Err>#[^:]+)))(:(?<C2>((?<C2ColDollar>\$)?(?<C2Col>[A-Z]+)(?<C2RowDollar>\$)?(?<C2Row>[0-9]+)|(?<C2Err>#[^:]+))))?$", RegexOptions.Compiled);
        private string _text;

        internal RangeComponents(string range)
        {
            this._text = range;
            this._m = _rgxRef.Match(range);
            if (!this._m.Success)
            {
                throw new ArgumentException("Invalid range: " + range);
            }
        }

        public string Cell1
        {
            get
            {
                return this._m.Groups["C1"].Value;
            }
        }

        private string Cell1Col
        {
            get
            {
                return this._m.Groups["C1Col"].Value;
            }
        }

        public string Cell1ColDollar
        {
            get
            {
                return this._m.Groups["C1ColDollar"].Value;
            }
        }

        public string Cell1Error
        {
            get
            {
                return this._m.Groups["C1Err"].Value;
            }
        }

        private string Cell1Row
        {
            get
            {
                return this._m.Groups["C1Row"].Value;
            }
        }

        public RowColumn Cell1RowColumn
        {
            get
            {
                return ExcelAddress.ToRowColumn(this.Cell1Col + this.Cell1Row);
            }
        }

        public string Cell1RowDollar
        {
            get
            {
                return this._m.Groups["C1RowDollar"].Value;
            }
        }

        public string Cell2
        {
            get
            {
                return this._m.Groups["C2"].Value;
            }
        }

        private string Cell2Col
        {
            get
            {
                return this._m.Groups["C2Col"].Value;
            }
        }

        public string Cell2ColDollar
        {
            get
            {
                return this._m.Groups["C2ColDollar"].Value;
            }
        }

        public string Cell2Error
        {
            get
            {
                return this._m.Groups["C2Err"].Value;
            }
        }

        private string Cell2Row
        {
            get
            {
                return this._m.Groups["C2Row"].Value;
            }
        }

        public RowColumn Cell2RowColumn
        {
            get
            {
                if (this.Cell2Col != "")
                {
                    return ExcelAddress.ToRowColumn(this.Cell2Col + this.Cell2Row);
                }
                return new RowColumn();
            }
        }

        public string Cell2RowDollar
        {
            get
            {
                return this._m.Groups["C2RowDollar"].Value;
            }
        }

        public string EscapedSheetName
        {
            get
            {
                string str = this._m.Groups["SheetName"].Value;
                if (str != "")
                {
                    return str;
                }
                return "";
            }
        }

        public string SheetName
        {
            get
            {
                string str = this._m.Groups["SheetName"].Value;
                if (!(str != ""))
                {
                    return "";
                }
                if (str.StartsWith("'") && str.EndsWith("'"))
                {
                    str = str.Replace("''", "'");
                    str = str.Substring(1, str.Length - 2);
                }
                return str;
            }
        }

        public string Text
        {
            get
            {
                return this._text;
            }
        }
    }
}

