namespace OpenExcel.OfficeOpenXml.Style
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelFill
    {
        private uint? _fillId;
        private IStylable _stylable;
        private DocumentStyles _styles;

        internal ExcelFill(IStylable stylable, DocumentStyles styles, uint? fillId)
        {
            this._stylable = stylable;
            this._styles = styles;
            this._fillId = fillId;
            if (this._fillId.HasValue)
            {
                this.FillObject = (Fill) this._styles.GetFill(this._fillId.Value).CloneNode(true);
            }
            else
            {
                this.FillObject = new Fill();
            }
        }

        private void EnsurePatternFill()
        {
            if (this.FillObject.GradientFill != null)
            {
                this.FillObject.GradientFill.Remove();
            }
            if (this.FillObject.PatternFill == null)
            {
                this.FillObject.PatternFill = new PatternFill();
            }
        }

        public string BackgroundColor
        {
            get
            {
                if (this.FillObject.PatternFill != null)
                {
                    return this.FillObject.PatternFill.BackgroundColor.Rgb.ToString();
                }
                return null;
            }
            set
            {
                this.EnsurePatternFill();
                this.FillObject.PatternFill.PatternType = (PatternValues)1;
                DocumentFormat.OpenXml.Spreadsheet.BackgroundColor color = new DocumentFormat.OpenXml.Spreadsheet.BackgroundColor {
                    Rgb = new HexBinaryValue(value)
                };
                this.FillObject.PatternFill.BackgroundColor = color;
                if (this._stylable != null)
                {
                    this._stylable.Style.Fill = this;
                }
            }
        }

        internal Fill FillObject { get; set; }

        public string ForegroundColor
        {
            get
            {
                if (this.FillObject.PatternFill != null)
                {
                    return this.FillObject.PatternFill.ForegroundColor.Rgb.ToString();
                }
                return null;
            }
            set
            {
                this.EnsurePatternFill();
                this.FillObject.PatternFill.PatternType = (PatternValues)1;
                DocumentFormat.OpenXml.Spreadsheet.ForegroundColor color = new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor {
                    Rgb = new HexBinaryValue(value)
                };
                this.FillObject.PatternFill.ForegroundColor = color;
                if (this._stylable != null)
                {
                    this._stylable.Style.Fill = this;
                }
            }
        }
    }
}

