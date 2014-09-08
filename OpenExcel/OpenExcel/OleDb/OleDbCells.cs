namespace OpenExcel.OleDb
{
    using OpenExcel.Common;
    using System;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class OleDbCells
    {
        public OleDbCells(OleDbExcelWorksheet wsheet)
        {
            this.Worksheet = wsheet;
        }

        public OleDbCell this[string address]
        {
            get
            {
                RowColumn column = ExcelAddress.ToRowColumn(address);
                return this[column.Row, column.Column];
            }
        }

        public OleDbCell this[uint row, uint col]
        {
            get
            {
                return new OleDbCell(row, col, this.Worksheet);
            }
        }

        public OleDbExcelWorksheet Worksheet { get; protected set; }
    }
}

