namespace OpenExcel.OleDb
{
    using OpenExcel.Common;
    using System;
    using System.Data;
    using System.Runtime.CompilerServices;

    public class OleDbCell
    {
        public OleDbCell(uint row, uint col, OleDbExcelWorksheet wsheet)
        {
            this.Row = row;
            this.Column = col;
            this.Worksheet = wsheet;
        }

        public string Address
        {
            get
            {
                return RowColumn.ToAddress(this.Row, this.Column);
            }
        }

        public uint Column { get; protected set; }

        public uint Row { get; protected set; }

        public object Value
        {
            get
            {
                Func<DataTable>[] funcArray = null;
                switch (this.Worksheet.Reader.Options.IMEX)
                {
                    case IMEXOptions.Yes:
                        funcArray = new Func<DataTable>[] { new Func<DataTable>(this.Worksheet.GetCachedTableIMEX) };
                        break;

                    case IMEXOptions.No:
                        funcArray = new Func<DataTable>[] { new Func<DataTable>(this.Worksheet.GetCachedTable) };
                        break;

                    case IMEXOptions.IMEXFirst:
                        funcArray = new Func<DataTable>[] { new Func<DataTable>(this.Worksheet.GetCachedTableIMEX), new Func<DataTable>(this.Worksheet.GetCachedTable) };
                        break;

                    case IMEXOptions.NoIMEXFirst:
                        funcArray = new Func<DataTable>[] { new Func<DataTable>(this.Worksheet.GetCachedTable), new Func<DataTable>(this.Worksheet.GetCachedTableIMEX) };
                        break;
                }
                foreach (Func<DataTable> func in funcArray)
                {
                    DataTable table = func();
                    if (this.Row > table.Rows.Count)
                    {
                        return null;
                    }
                    if (this.Column > table.Columns.Count)
                    {
                        return null;
                    }
                    object obj2 = table.Rows[((int) this.Row) - 1][((int) this.Column) - 1];
                    if (obj2 != DBNull.Value)
                    {
                        return obj2;
                    }
                }
                return DBNull.Value;
            }
        }

        public OleDbExcelWorksheet Worksheet { get; protected set; }
    }
}

