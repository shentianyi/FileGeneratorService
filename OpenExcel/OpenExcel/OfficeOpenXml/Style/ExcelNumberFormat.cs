namespace OpenExcel.OfficeOpenXml.Style
{
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.CompilerServices;

    public class ExcelNumberFormat
    {
        private static Dictionary<uint, string> _builtInFormats_Global;
        private IStylable _stylable;
        private DocumentStyles _styles;

        static ExcelNumberFormat()
        {
            Dictionary<uint, string> dictionary = new Dictionary<uint, string>();
            dictionary.Add(0, "General");
            dictionary.Add(1, "0");
            dictionary.Add(2, "0.00");
            dictionary.Add(3, "#,##0");
            dictionary.Add(4, "#,##0.00");
            dictionary.Add(9, "0%");
            dictionary.Add(10, "0.00%");
            dictionary.Add(11, "0.00E+00");
            dictionary.Add(12, "# ?/?");
            dictionary.Add(13, "# ??/??");
            dictionary.Add(14, "m/d/yyyy");
            dictionary.Add(15, "d-mmm-yy");
            dictionary.Add(0x10, "d-mmm");
            dictionary.Add(0x11, "mmm-yy");
            dictionary.Add(0x12, "h:mm AM/PM");
            dictionary.Add(0x13, "h:mm:ss AM/PM");
            dictionary.Add(20, "h:mm");
            dictionary.Add(0x15, "h:mm:ss");
            dictionary.Add(0x16, "m/d/yy h:mm");
            dictionary.Add(0x25, "#,##0 ;(#,##0)");
            dictionary.Add(0x26, "#,##0 ;[Red](#,##0)");
            dictionary.Add(0x27, "#,##0.00;(#,##0.00)");
            dictionary.Add(40, "#,##0.00;[Red](#,##0.00)");
            dictionary.Add(0x2d, "mm:ss");
            dictionary.Add(0x2e, "[h]:mm:ss");
            dictionary.Add(0x2f, "mmss.0");
            dictionary.Add(0x30, "##0.0E+0");
            dictionary.Add(0x31, "@");
            _builtInFormats_Global = dictionary;
        }

        internal ExcelNumberFormat(IStylable stylable, DocumentStyles styles, uint numFmtId)
        {
            this._stylable = stylable;
            this._styles = styles;
            this.NumFmtId = numFmtId;
        }

        public string Format
        {
            get
            {
                if (_builtInFormats_Global.ContainsKey(this.NumFmtId))
                {
                    return _builtInFormats_Global[this.NumFmtId];
                }
                return (string) this._styles.GetNumberingFormat(this.NumFmtId).FormatCode;
            }
            set
            {
                uint key;
                KeyValuePair<uint, string> pair = (from i in _builtInFormats_Global
                    where i.Value == value
                    select i).FirstOrDefault<KeyValuePair<uint, string>>();
                if (pair.Value == value)
                {
                    key = pair.Key;
                }
                else
                {
                    NumberingFormat nfNew = new NumberingFormat {
                        FormatCode = value
                    };
                    key = this._styles.EnsureCustomNumberingFormat(nfNew);
                }
                if (key != this.NumFmtId)
                {
                    this.NumFmtId = key;
                    if (this._stylable != null)
                    {
                        this._stylable.Style.NumberFormat = this;
                    }
                }
            }
        }

        internal uint NumFmtId { get; set; }
    }
}

