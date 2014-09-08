namespace OpenExcel.OfficeOpenXml.Style
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelBorder
    {
        private uint? _borderId;
        private IStylable _stylable;
        private DocumentStyles _styles;

        public ExcelBorder(IStylable stylable, DocumentStyles styles, uint? borderId)
        {
            this._stylable = stylable;
            this._styles = styles;
            this._borderId = borderId;
            if (this._borderId.HasValue)
            {
                this.BorderObject = (Border) this._styles.GetBorder(this._borderId.Value).CloneNode(true);
            }
            else
            {
                this.BorderObject = new Border();
            }
        }

        private Color GetBorderColor(BorderPropertiesType b)
        {
            return b.Color;
        }

        private ExcelBorderStyleValues GetBorderStyle(BorderPropertiesType b)
        {
            return (ExcelBorderStyleValues)b.Style.Value;
        }

        private void SetBorderColor(BorderPropertiesType b, Color val)
        {
            b.Color = val;
            if (this._stylable != null)
            {
                this._stylable.Style.Border = this;
            }
        }

        private void SetBorderStyle(BorderPropertiesType b, ExcelBorderStyleValues val)
        {
            b.Style = (EnumValue<BorderStyleValues>)((BorderStyleValues)(int)val);
            if (this._stylable != null)
            {
                this._stylable.Style.Border = this;
            }
        }

        internal Border BorderObject { get; set; }

        public string BottomColor
        {
            get
            {
                if (this.GetBorderColor(this.BorderObject.BottomBorder) == null)
                {
                    return "";
                }
                return this.GetBorderColor(this.BorderObject.BottomBorder).Rgb.Value;
            }
            set
            {
                Color val = new Color {
                    Rgb = value
                };
                this.SetBorderColor(this.BorderObject.BottomBorder, val);
            }
        }

        public ExcelBorderStyleValues BottomStyle
        {
            get
            {
                return this.GetBorderStyle(this.BorderObject.BottomBorder);
            }
            set
            {
                this.SetBorderStyle(this.BorderObject.BottomBorder, value);
            }
        }

        public string LeftColor
        {
            get
            {
                if (this.GetBorderColor(this.BorderObject.LeftBorder) == null)
                {
                    return "";
                }
                return this.GetBorderColor(this.BorderObject.LeftBorder).Rgb.Value;
            }
            set
            {
                Color val = new Color {
                    Rgb = value
                };
                this.SetBorderColor(this.BorderObject.LeftBorder, val);
            }
        }

        public ExcelBorderStyleValues LeftStyle
        {
            get
            {
                return this.GetBorderStyle(this.BorderObject.LeftBorder);
            }
            set
            {
                this.SetBorderStyle(this.BorderObject.LeftBorder, value);
            }
        }

        public string RightColor
        {
            get
            {
                if (this.GetBorderColor(this.BorderObject.RightBorder) == null)
                {
                    return "";
                }
                return this.GetBorderColor(this.BorderObject.RightBorder).Rgb.Value;
            }
            set
            {
                Color val = new Color {
                    Rgb = value
                };
                this.SetBorderColor(this.BorderObject.RightBorder, val);
            }
        }

        public ExcelBorderStyleValues RightStyle
        {
            get
            {
                return this.GetBorderStyle(this.BorderObject.RightBorder);
            }
            set
            {
                this.SetBorderStyle(this.BorderObject.RightBorder, value);
            }
        }

        public string TopColor
        {
            get
            {
                if (this.GetBorderColor(this.BorderObject.TopBorder) == null)
                {
                    return "";
                }
                return this.GetBorderColor(this.BorderObject.TopBorder).Rgb.Value;
            }
            set
            {
                Color val = new Color {
                    Rgb = value
                };
                this.SetBorderColor(this.BorderObject.TopBorder, val);
            }
        }

        public ExcelBorderStyleValues TopStyle
        {
            get
            {
                return this.GetBorderStyle(this.BorderObject.TopBorder);
            }
            set
            {
                this.SetBorderStyle(this.BorderObject.TopBorder, value);
            }
        }
    }
}

