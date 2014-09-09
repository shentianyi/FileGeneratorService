namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.Common;
    using OpenExcel.OfficeOpenXml.Internal;
    using OpenExcel.OfficeOpenXml.Style;
    using OpenExcel.Utilities;
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelCell : IStylable
    {
        internal ExcelCell(uint row, uint col, ExcelWorksheet wsheet,bool _isHead=false)
        {
            this.Row = row;
            this.Column = col;
            this.Worksheet = wsheet;
            this.IsHead = _isHead;
        }

        private object GetValue()
        {
            CellProxy cell = this.Worksheet.GetCell(this.Row, this.Column);
            if (cell != null)
            {
                if (cell.DataType.HasValue)
                {
                    CellValues values = cell.DataType.Value;
                    switch (values)
                    {
                        case CellValues.Number:
                            return cell.Value;

                        case CellValues.InlineString:
                            return cell.Value;
                    }
                    if ((values == CellValues.SharedString) && (cell.Value != null))
                    {
                        return this.Worksheet.Document.SharedStrings.Get(Convert.ToUInt32(cell.Value));
                    }
                }
                if (cell.StyleIndex.HasValue)
                {
                    CellFormat cellFormat = this.Worksheet.Document.Styles.GetCellFormat(cell.StyleIndex.Value);
                    if (!this.Worksheet.Document.Styles.IsDateFormat(cellFormat))
                    {
                        if (cell.Value != null)
                        {
                            return cell.Value;
                        }
                    }
                    else if (cell.Value != null)
                    {
                        return DateTime.FromOADate((double) cell.Value);
                    }
                }
                else if (cell.Value != null)
                {
                    return cell.Value;
                }
            }
            return null;
        }

        private void SetValue(object value)
        {
            CellProxy proxy = this.Worksheet.EnsureCell(this.Row, this.Column);
            bool flag = false;
            if (proxy.StyleIndex.HasValue)
            {
                CellFormat cellFormat = this.Worksheet.Document.Styles.GetCellFormat(proxy.StyleIndex.Value);
                if (this.Worksheet.Document.Styles.IsDateFormat(cellFormat))
                {
                    flag = true;
                }
            }
            if (value == null)
            {
                proxy.DataType = null;
                proxy.Value = null;
            }
            else
            {
                Type valueType = value.GetType();
                if (valueType == typeof(DateTime))
                {
                    proxy.DataType = null;
                    if (!flag)
                    {
                        CellFormat cfNew = new CellFormat {
                            ApplyNumberFormat = true,
                            NumberFormatId = 14
                        };
                        uint? styleIndex = proxy.StyleIndex;
                        uint num = this.Worksheet.Document.Styles.MergeAndRegisterCellFormat(cfNew, styleIndex.HasValue ? ((UInt32Value) styleIndex.GetValueOrDefault()) : null, false);
                        proxy.StyleIndex = new uint?(num);
                    }
                    proxy.Value = ((DateTime) value).ToOADate();
                }
                else if (ValueChecker.IsNumeric(valueType))
                {
                    if (flag)
                    {
                        CellFormat format4 = new CellFormat {
                            NumberFormatId = 0
                        };
                        uint? nullable6 = proxy.StyleIndex;
                        uint num2 = this.Worksheet.Document.Styles.MergeAndRegisterCellFormat(format4, nullable6.HasValue ? ((UInt32Value) nullable6.GetValueOrDefault()) : null, false);
                        proxy.StyleIndex = new uint?(num2);
                    }
                    proxy.DataType = (CellValues)1;
                    proxy.Value = value;
                }
                else
                {
                    if (flag)
                    {
                        CellFormat format6 = new CellFormat {
                            NumberFormatId = 0
                        };
                        uint? nullable7 = proxy.StyleIndex;
                        uint num3 = this.Worksheet.Document.Styles.MergeAndRegisterCellFormat(format6, nullable7.HasValue ? ((UInt32Value) nullable7.GetValueOrDefault()) : null, false);
                        proxy.StyleIndex = new uint?(num3);
                    }
                    string valueStr = value.ToString();
                    int num4 = this.Worksheet.Document.SharedStrings.Put(valueStr);
                    proxy.DataType = (CellValues)3;
                    proxy.Value = (uint) num4;
                }
            }
        }

        public string Address
        {
            get
            {
                return RowColumn.ToAddress(this.Row, this.Column);
            }
        }

        public uint Column { get; protected set; }

        public ExcelCellFormula Formula
        {
            get
            {
                return new ExcelCellFormula(this.Row, this.Column, this.Worksheet);
            }
        }

        public uint Row { get; protected set; }
        private bool isHead = false;
        public bool IsHead { get { return isHead; } set { isHead = value; } }

        public ExcelStyle Style
        {
            get
            {
                uint? baseStyleIndex = null;
                CellProxy cell = this.Worksheet.GetCell(this.Row, this.Column);
                if (cell != null)
                {
                    baseStyleIndex = cell.StyleIndex;
                }
                return new ExcelStyle(this, this.Worksheet.Document.Styles, baseStyleIndex,this.IsHead);
            }
            set
            {
                if (value != null)
                {
                    CellProxy proxy = this.Worksheet.EnsureCell(this.Row, this.Column);
                    uint? styleIndex = value.StyleIndex;
                    CellFormat cellFormat = this.Worksheet.Document.Styles.GetCellFormat(styleIndex.HasValue ? styleIndex.GetValueOrDefault() : 0);
                    uint? nullable2 = proxy.StyleIndex;
                    proxy.StyleIndex = new uint?(this.Worksheet.Document.Styles.MergeAndRegisterCellFormat(cellFormat, nullable2.HasValue ? ((UInt32Value) nullable2.GetValueOrDefault()) : null, false));
                    this.Worksheet.Modified = true;
                }
                else
                {
                    CellProxy cell = this.Worksheet.GetCell(this.Row, this.Column);
                    if (cell != null)
                    {
                        cell.StyleIndex = null;
                        this.Worksheet.Modified = true;
                    }
                }
            }
        }

        public object Value
        {
            get
            {
                return this.GetValue();
            }
            set
            {
                this.SetValue(value);
            }
        }

        public ExcelWorksheet Worksheet { get; protected set; }
    }
}

