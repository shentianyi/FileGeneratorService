namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.OfficeOpenXml.Style;
    using System;

    public class ExcelRow : IStylable
    {
        private uint _row;
        private ExcelWorksheet _wsheet;

        internal ExcelRow(uint row, ExcelWorksheet wsheet)
        {
            this._row = row;
            this._wsheet = wsheet;
        }

        public double? Height
        {
            get
            {
                DocumentFormat.OpenXml.Spreadsheet.Row row = this._wsheet.GetRow(this._row);
                if (row != null)
                {
                    return new double?((double) row.Height);
                }
                return null;
            }
            set
            {
                if (value.HasValue)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row row = this._wsheet.EnsureRow(this._row);
                    row.Height = value.Value;
                    row.CustomHeight = true;
                }
                else
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row row2 = this._wsheet.EnsureRow(this._row);
                    row2.Height = null;
                    row2.CustomHeight = null;
                }
                this._wsheet.Modified = true;
            }
        }

        public bool Hidden
        {
            get
            {
                DocumentFormat.OpenXml.Spreadsheet.Row row = this._wsheet.GetRow(this._row);
                if ((row != null) && row.Hidden.HasValue)
                {
                    return row.Hidden.Value;
                }
                return false;
            }
            set
            {
                this._wsheet.EnsureRow(this._row).Hidden = value;
                this._wsheet.Modified = true;
            }
        }

        public uint Row
        {
            get
            {
                return this._row;
            }
        }

        public ExcelStyle Style
        {
            get
            {
                uint? baseStyleIndex = null;
                DocumentFormat.OpenXml.Spreadsheet.Row row = this._wsheet.GetRow(this._row);
                if ((row != null) && (row.StyleIndex != null))
                {
                    baseStyleIndex = new uint?((uint) row.StyleIndex);
                }
                return new ExcelStyle(this, this._wsheet.Document.Styles, baseStyleIndex);
            }
            set
            {
                if (value != null)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row row = this._wsheet.EnsureRow(this._row);
                    uint? styleIndex = value.StyleIndex;
                    CellFormat cellFormat = this._wsheet.Document.Styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0);
                    row.StyleIndex = this._wsheet.Document.Styles.MergeAndRegisterCellFormat(cellFormat, row.StyleIndex, false);
                    row.CustomFormat = true;
                    this._wsheet.Modified = true;
                }
                else
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row row2 = this._wsheet.GetRow(this._row);
                    if (row2 != null)
                    {
                        row2.StyleIndex = null;
                        row2.CustomFormat = false;
                        this._wsheet.Modified = true;
                    }
                }
            }
        }
    }
}

