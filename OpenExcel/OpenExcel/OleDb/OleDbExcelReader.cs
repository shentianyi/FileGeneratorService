namespace OpenExcel.OleDb
{
    using System;
    using System.Data.OleDb;
    using System.Runtime.CompilerServices;

    public class OleDbExcelReader : IDisposable
    {
        private OleDbConnection _conn;
        private static string _connStrIMEX = ("Provider=" + _provider + ";Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;ReadOnly=True;IMEX=1\"");
        private static string _connStrNoIMEX = ("Provider=" + _provider + ";Data Source={0};Extended Properties=\"Excel 8.0;HDR=No;ReadOnly=True;\"");
        private bool _disposed;
        private string _path;
        private static string _provider = "Microsoft.ACE.OLEDB.12.0";

        public OleDbExcelReader(string path)
        {
            this.Worksheets = new OleDbExcelWorksheets(this);
            this._path = path;
            this.Options = new ReaderOptions();
        }

        public OleDbExcelReader(string path, ReaderOptions options)
        {
            this.Worksheets = new OleDbExcelWorksheets(this);
            this._path = path;
            this.Options = options;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (this._conn != null)
                {
                    this._conn.Dispose();
                }
                this._disposed = true;
            }
        }

        ~OleDbExcelReader()
        {
            this.Dispose(false);
        }

        internal OleDbConnection OpenConnection(bool useImex)
        {
            if (useImex)
            {
                this._conn = new OleDbConnection(string.Format(_connStrIMEX, this._path));
            }
            else
            {
                this._conn = new OleDbConnection(string.Format(_connStrNoIMEX, this._path));
            }
            this._conn.Open();
            return this._conn;
        }

        public ReaderOptions Options { get; protected set; }

        public OleDbExcelWorksheets Worksheets { get; protected set; }
    }
}

