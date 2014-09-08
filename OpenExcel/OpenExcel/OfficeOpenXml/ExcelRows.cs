namespace OpenExcel.OfficeOpenXml
{
    using OpenExcel.Common;
    using System;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class ExcelRows
    {
        internal ExcelRows(ExcelWorksheet wsheet)
        {
            this.Worksheet = wsheet;
        }

        public ExcelRow this[uint row]
        {
            get
            {
                if ((row < 1) || (row > ExcelConstraints.MaxRows))
                {
                    throw new ArgumentException("Invalid row value");
                }
                return new ExcelRow(row, this.Worksheet);
            }
        }

        public ExcelWorksheet Worksheet { get; protected set; }
    }
}

