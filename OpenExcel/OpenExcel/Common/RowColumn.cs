namespace OpenExcel.Common
{
    using System;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential)]
    public struct RowColumn
    {
        public uint Row { get; set; }
        public uint Column { get; set; }
        public static string ToAddress(uint row, uint col)
        {
            if ((row < 1) || (row > ExcelConstraints.MaxRows))
            {
                throw new ArgumentException("Invalid row value");
            }
            if ((col < 1) || (col > ExcelConstraints.MaxColumns))
            {
                throw new ArgumentException("Invalid column value");
            }
            return (ExcelAddress.ColumnIndexToName(col) + row);
        }

        public string ToAddress()
        {
            if ((this.Row < 1) || (this.Row > ExcelConstraints.MaxRows))
            {
                throw new ArgumentException("Invalid row value");
            }
            if ((this.Column < 1) || (this.Column > ExcelConstraints.MaxColumns))
            {
                throw new ArgumentException("Invalid column value");
            }
            return (ExcelAddress.ColumnIndexToName(this.Column) + this.Row);
        }
    }
}

