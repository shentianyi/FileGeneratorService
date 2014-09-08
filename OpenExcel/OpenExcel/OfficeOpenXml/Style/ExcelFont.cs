namespace OpenExcel.OfficeOpenXml.Style
{
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelFont
    {
        private uint? _fontId;
        private IStylable _stylable;
        private DocumentStyles _styles;

        internal ExcelFont(IStylable stylable, DocumentStyles styles, uint? fontId)
        {
            this._stylable = stylable;
            this._styles = styles;
            this._fontId = fontId;
            if (this._fontId.HasValue)
            {
                this.FontObject = (Font) this._styles.GetFont(this._fontId.Value).CloneNode(true);
            }
            else
            {
                this.FontObject = new Font();
            }
        }

        public bool Bold
        {
            get
            {
                return (bool) this.FontObject.Bold.Val;
            }
            set
            {
                if (this.FontObject.Bold == null)
                {
                    this.FontObject.Bold = new DocumentFormat.OpenXml.Spreadsheet.Bold();
                }
                this.FontObject.Bold.Val = value;
                if (this._stylable != null)
                {
                    this._stylable.Style.Font = this;
                }
            }
        }

        public string Color
        {
            get
            {
                return this.FontObject.Color.Rgb.Value;
            }
            set
            {
                DocumentFormat.OpenXml.Spreadsheet.Color color = new DocumentFormat.OpenXml.Spreadsheet.Color {
                    Rgb = value
                };
                this.FontObject.Color = color;
                if (this._stylable != null)
                {
                    this._stylable.Style.Font = this;
                }
            }
        }

        internal Font FontObject { get; set; }

        public bool Italic
        {
            get
            {
                return (bool) this.FontObject.Italic.Val;
            }
            set
            {
                if (this.FontObject.Italic == null)
                {
                    this.FontObject.Italic = new DocumentFormat.OpenXml.Spreadsheet.Italic();
                }
                this.FontObject.Italic.Val = value;
                if (this._stylable != null)
                {
                    this._stylable.Style.Font = this;
                }
            }
        }

        public string Name
        {
            get
            {
                return (string) this.FontObject.FontName.Val;
            }
            set
            {
                if (this.FontObject.FontName == null)
                {
                    this.FontObject.FontName = new FontName();
                }
                this.FontObject.FontName.Val = value;
                if (this.FontObject.FontScheme == null)
                {
                    this.FontObject.FontScheme = new FontScheme();
                }
                this.FontObject.FontScheme.Val = (FontSchemeValues)0;
                if (this._stylable != null)
                {
                    this._stylable.Style.Font = this;
                }
            }
        }

        public double Size
        {
            get
            {
                return this.FontObject.FontSize.Val.Value;
            }
            set
            {
                if (this.FontObject.FontSize == null)
                {
                    this.FontObject.FontSize = new FontSize();
                }
                this.FontObject.FontSize.Val = value;
                if (this._stylable != null)
                {
                    this._stylable.Style.Font = this;
                }
            }
        }
    }
}

