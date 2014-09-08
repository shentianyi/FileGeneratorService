namespace OpenExcel.OleDb
{
    using System;
    using System.Data;
    using System.Data.OleDb;
    using System.Runtime.CompilerServices;

    public class OleDbExcelWorksheet
    {
        private WeakReference _cachedTable;
        private WeakReference _cachedTableIMEX;

        internal OleDbExcelWorksheet(string name, OleDbExcelReader parent)
        {
            this.Name = name;
            this.Reader = parent;
            this.Cells = new OleDbCells(this);
        }

        internal DataTable GetCachedTable()
        {
            DataTable table;
            if ((this._cachedTable == null) || ((table = (DataTable) this._cachedTable.Target) == null))
            {
                using (OleDbConnection connection = this.Reader.OpenConnection(false))
                {
                    OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM ['" + this.Name.Replace("'", "''") + "$']", connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand);
                    table = new DataTable();
                    adapter.Fill(table);
                    this._cachedTable = new WeakReference(table);
                }
            }
            return table;
        }

        internal DataTable GetCachedTableIMEX()
        {
            DataTable table;
            if ((this._cachedTableIMEX == null) || ((table = (DataTable) this._cachedTableIMEX.Target) == null))
            {
                using (OleDbConnection connection = this.Reader.OpenConnection(true))
                {
                    OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM ['" + this.Name.Replace("'", "''") + "$']", connection);
                    OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand);
                    table = new DataTable();
                    adapter.Fill(table);
                    this._cachedTableIMEX = new WeakReference(table);
                }
            }
            return table;
        }

        public OleDbCells Cells { get; protected set; }

        public string Name { get; protected set; }

        public OleDbExcelReader Reader { get; protected set; }
    }
}

