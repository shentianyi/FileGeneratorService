namespace OpenExcel.Common
{
    using System;
    using System.Runtime.CompilerServices;

    public class SheetChange
    {
        public int ColumnDelta { get; set; }

        public uint ColumnStart { get; set; }

        public int RowDelta { get; set; }

        public uint RowStart { get; set; }

        public string SheetName { get; set; }
    }
}

