namespace OpenExcel.OfficeOpenXml
{
    using System;
    using System.Runtime.CompilerServices;

    public class ExcelWorkbook
    {
        internal ExcelWorkbook(ExcelDocument parent)
        {
            this.Document = parent;
            this.Worksheets = new ExcelWorksheets(parent);
        }

        public ExcelDocument Document { get; protected set; }

        public ExcelWorksheets Worksheets { get; protected set; }
    }
}

