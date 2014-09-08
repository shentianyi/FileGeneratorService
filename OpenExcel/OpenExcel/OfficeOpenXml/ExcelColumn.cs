namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Style;
    using System;

    public class ExcelColumn : IStylable
    {
        private uint _col;
        private ExcelWorksheet _wsheet;

        internal ExcelColumn(uint col, ExcelWorksheet wsheet)
        {
            this._col = col;
            this._wsheet = wsheet;
        }

        public uint Column
        {
            get
            {
                return this._col;
            }
        }

        public bool Hidden
        {
            get
            {
                DocumentFormat.OpenXml.Spreadsheet.Column columnDefinition = this._wsheet.GetColumnDefinition(this._col);
                if ((columnDefinition != null) && columnDefinition.Hidden.HasValue)
                {
                    return columnDefinition.Hidden.Value;
                }
                return false;
            }
            set
            {
                this._wsheet.EnsureColumnDefinition(this._col).Hidden = value;
                this._wsheet.Modified = true;
            }
        }

        public ExcelStyle Style
        {
            get
            {
                uint? baseStyleIndex = null;
                DocumentFormat.OpenXml.Spreadsheet.Column columnDefinition = this._wsheet.GetColumnDefinition(this._col);
                if (columnDefinition != null)
                {
                    baseStyleIndex = new uint?((uint) columnDefinition.Style);
                }
                return new ExcelStyle(this, this._wsheet.Document.Styles, baseStyleIndex);
            }
            set
            {
                if (value != null)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column column = this._wsheet.EnsureColumnDefinition(this._col);
                    uint? styleIndex = value.StyleIndex;
                    CellFormat cellFormat = this._wsheet.Document.Styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0);
                    column.Style = this._wsheet.Document.Styles.MergeAndRegisterCellFormat(cellFormat, column.Style, false);
                    this._wsheet.Modified = true;
                }
                else
                {
                    this._wsheet.EnsureColumnDefinition(this._col).Style = null;
                    this._wsheet.Modified = true;
                }
            }
        }

        public double? Width
        {
            get
            {
                DocumentFormat.OpenXml.Spreadsheet.Column columnDefinition = this._wsheet.GetColumnDefinition(this._col);
                if (columnDefinition != null)
                {
                    return new double?((double) columnDefinition.Width);
                }
                return null;
            }
            set
            {
                if (value.HasValue)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Column column = this._wsheet.EnsureColumnDefinition(this._col);
                    column.Width = value.Value;
                    column.CustomWidth = true;
                }
                else
                {
                    this._wsheet.EnsureColumnDefinition(this._col);
                    this._wsheet.DeleteColumnDefinition(this._col);
                }
                this._wsheet.Modified = true;
            }
        }
    }
}

