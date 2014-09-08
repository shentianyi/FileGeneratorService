namespace OpenExcel.OfficeOpenXml.Style
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelStyle
    {
        private IStylable _stylable;
        private DocumentStyles _styles;

        internal ExcelStyle(IStylable stylable, DocumentStyles styles, uint? baseStyleIndex)
        {
            this._stylable = stylable;
            this._styles = styles;
            this.StyleIndex = baseStyleIndex;
        }

        public void ApplySettings(DocumentFormat.OpenXml.Spreadsheet.Font font, DocumentFormat.OpenXml.Spreadsheet.Fill fill, params ExcelBorder[] borders)
        {
        }

        public uint GetBorderId()
        {
            uint? styleIndex = this.StyleIndex;
            return (this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).BorderId ?? 0);
        }

        public ExcelBorder Border
        {
            get
            {
                uint? styleIndex = this.StyleIndex;
                uint num = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).BorderId ?? 0;
                return new ExcelBorder(this._stylable, this._styles, new uint?(num));
            }
            set
            {
                uint? styleIndex = this.StyleIndex;
                uint baseBordersIdx = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).BorderId ?? 0;
                uint num2 = this._styles.MergeAndRegisterBorder(value.BorderObject, baseBordersIdx, false);
                if (num2 != baseBordersIdx)
                {
                    CellFormat cfNew = new CellFormat {
                        BorderId = num2,
                        ApplyBorder = true
                    };
                    uint? nullable2 = this.StyleIndex;
                    this.StyleIndex = new uint?(this._styles.MergeAndRegisterCellFormat(cfNew, nullable2.HasValue ? ((UInt32Value) nullable2.GetValueOrDefault()) : null, false));
                    if (this._stylable != null)
                    {
                        this._stylable.Style = this;
                    }
                }
            }
        }

        public ExcelFill Fill
        {
            get
            {
                uint? styleIndex = this.StyleIndex;
                uint num = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).FillId ?? 0;
                return new ExcelFill(this._stylable, this._styles, new uint?(num));
            }
            set
            {
                uint? styleIndex = this.StyleIndex;
                uint baseFillsIdx = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).FillId ?? 0;
                uint num2 = this._styles.MergeAndRegisterFill(value.FillObject, baseFillsIdx, false);
                if (num2 != baseFillsIdx)
                {
                    CellFormat cfNew = new CellFormat {
                        FillId = num2,
                        ApplyFill = true
                    };
                    uint? nullable2 = this.StyleIndex;
                    this.StyleIndex = new uint?(this._styles.MergeAndRegisterCellFormat(cfNew, nullable2.HasValue ? ((UInt32Value) nullable2.GetValueOrDefault()) : null, false));
                    if (this._stylable != null)
                    {
                        this._stylable.Style = this;
                    }
                }
            }
        }

        public ExcelFont Font
        {
            get
            {
                uint? styleIndex = this.StyleIndex;
                uint num = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).FontId ?? 0;
                return new ExcelFont(this._stylable, this._styles, new uint?(num));
            }
            set
            {
                uint? styleIndex = this.StyleIndex;
                uint baseFontsIdx = this._styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0).FontId ?? 0;
                uint num2 = this._styles.MergeAndRegisterFont(value.FontObject, baseFontsIdx, false);
                if (num2 != baseFontsIdx)
                {
                    CellFormat cfNew = new CellFormat {
                        FontId = num2,
                        ApplyFont = true
                    };
                    uint? nullable2 = this.StyleIndex;
                    this.StyleIndex = new uint?(this._styles.MergeAndRegisterCellFormat(cfNew, nullable2.HasValue ? ((UInt32Value) nullable2.GetValueOrDefault()) : null, false));
                    if (this._stylable != null)
                    {
                        this._stylable.Style = this;
                    }
                }
            }
        }

        public ExcelNumberFormat NumberFormat
        {
            get
            {
                if (this.StyleIndex.HasValue)
                {
                    CellFormat cellFormat = this._styles.GetCellFormat(this.StyleIndex.Value);
                    if (cellFormat.NumberFormatId != null)
                    {
                        return new ExcelNumberFormat(this._stylable, this._styles, (uint) cellFormat.NumberFormatId);
                    }
                }
                return new ExcelNumberFormat(this._stylable, this._styles, 0);
            }
            set
            {
                CellFormat cfNew = new CellFormat {
                    NumberFormatId = value.NumFmtId,
                    ApplyNumberFormat = true
                };
                uint? styleIndex = this.StyleIndex;
                this.StyleIndex = new uint?(this._styles.MergeAndRegisterCellFormat(cfNew, styleIndex.HasValue ? ((UInt32Value) styleIndex.GetValueOrDefault()) : null, false));
                if (this._stylable != null)
                {
                    this._stylable.Style = this;
                }
            }
        }

        internal uint? StyleIndex { get; set; }
    }
}

