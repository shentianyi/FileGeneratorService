namespace OpenExcel.OfficeOpenXml
{
    using OpenExcel.Common;
    using System;
    using System.Reflection;
    using System.Runtime.CompilerServices;

    public class ExcelColumns
    {
        internal ExcelColumns(ExcelWorksheet wsheet)
        {
            this.Worksheet = wsheet;
        }

        public ExcelColumn this[string colName]
        {
            get
            {
                return this[ExcelAddress.ColumnNameToIndex(colName)];
            }
        }

        public ExcelColumn this[uint col]
        {
            get
            {
                if ((col < 1) || (col > ExcelConstraints.MaxColumns))
                {
                    throw new ArgumentException("Invalid column value");
                }
                return new ExcelColumn(col, this.Worksheet);
            }
        }

        public ExcelWorksheet Worksheet { get; protected set; }
    }
}

