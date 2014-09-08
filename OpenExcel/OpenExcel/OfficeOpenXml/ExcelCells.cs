namespace OpenExcel.OfficeOpenXml
{
    using OpenExcel.Common;
    using System;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class ExcelCells
    {
        internal ExcelCells(ExcelWorksheet wsheet)
        {
            this.Worksheet = wsheet;
        }

        public ExcelCell this[string address]
        {
            get
            {
                RowColumn column = ExcelAddress.ToRowColumn(address);
                return this[column.Row, column.Column];
            }
        }

        public ExcelCell this[uint row, uint col]
        {
            get
            {
                if ((row < 1) || (row > ExcelConstraints.MaxRows))
                {
                    throw new ArgumentException("Invalid row value");
                }
                if ((col < 1) || (col > ExcelConstraints.MaxColumns))
                {
                    throw new ArgumentException("Invalid column value");
                }
                return new ExcelCell(row, col, this.Worksheet);
            }
        }

        public ExcelWorksheet Worksheet { get; protected set; }
    }
}

