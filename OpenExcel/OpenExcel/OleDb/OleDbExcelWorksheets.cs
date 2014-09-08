namespace OpenExcel.OleDb
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.Diagnostics;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Threading;

    public class OleDbExcelWorksheets : IEnumerable<OleDbExcelWorksheet>, IEnumerable
    {
        internal OleDbExcelWorksheets(OleDbExcelReader parent)
        {
            this.Reader = parent;
        }

        private IEnumerable<OleDbExcelWorksheet> EnumerateWorksheets()
        {
            using (OleDbConnection iteratorVariable0 = this.Reader.OpenConnection(false))
            {
                DataTable oleDbSchemaTable = iteratorVariable0.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                IEnumerable<string> iteratorVariable2 = from r in oleDbSchemaTable.Rows.Cast<DataRow>()
                    let tableName = (string) r["TABLE_NAME"]
                    where tableName.EndsWith("$") || (tableName.StartsWith("'") && tableName.EndsWith("$'"))
                    select tableName;
                foreach (string iteratorVariable3 in iteratorVariable2)
                {
                    string name = iteratorVariable3;
                    if (iteratorVariable3.StartsWith("'") && iteratorVariable3.EndsWith("$'"))
                    {
                        name = iteratorVariable3.Substring(1, iteratorVariable3.Length - 3).Replace("''", "'");
                    }
                    else
                    {
                        name = iteratorVariable3.Substring(0, iteratorVariable3.Length - 1);
                    }
                    yield return new OleDbExcelWorksheet(name, this.Reader);
                }
            }
        }

        public IEnumerator<OleDbExcelWorksheet> GetEnumerator()
        {
            foreach (OleDbExcelWorksheet iteratorVariable0 in this.EnumerateWorksheets())
            {
                yield return iteratorVariable0;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (OleDbExcelWorksheet iteratorVariable0 in this.EnumerateWorksheets())
            {
                yield return iteratorVariable0;
            }
        }

        public OleDbExcelWorksheet this[string name]
        {
            get
            {
                return new OleDbExcelWorksheet(name, this.Reader);
            }
        }

        public OleDbExcelReader Reader { get; protected set; }



    }
}

