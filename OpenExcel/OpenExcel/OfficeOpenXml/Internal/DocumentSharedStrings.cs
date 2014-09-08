namespace OpenExcel.OfficeOpenXml.Internal
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections.Generic;

    internal class DocumentSharedStrings
    {
        private bool _changed;
        private Dictionary<string, uint> _indexLookup;
        private SharedStringTablePart _ssPart;
        private SortedList<uint, string> _stringCache;
        private WorkbookPart _wpart;

        public DocumentSharedStrings(WorkbookPart wpart)
        {
            this._wpart = wpart;
            this._indexLookup = new Dictionary<string, uint>();
            this._ssPart = this._wpart.SharedStringTablePart;
            if (this._ssPart != null)
            {
                SharedStringTable sharedStringTable = this._ssPart.SharedStringTable;
                uint num = 0;
                foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>())
                {
                    string text = item.Text.Text;
                    this._indexLookup[text] = num;
                    num++;
                }
            }
        }

        private SharedStringTablePart EnsureSharedStringTablePart()
        {
            if (this._ssPart == null)
            {
                this._ssPart = this._wpart.AddNewPart<SharedStringTablePart>();
                this._ssPart.SharedStringTable = new SharedStringTable();
                this._ssPart.SharedStringTable.Save();
            }
            return this._ssPart;
        }

        public string Get(uint idx)
        {
            this.StringCacheLazyInit();
            return this._stringCache[idx];
        }

        public int Put(string valueStr)
        {
            uint num = 0;
            if (this._indexLookup.TryGetValue(valueStr, out num))
            {
                return (int) num;
            }
            uint count = (uint) this._indexLookup.Count;
            if (this._stringCache != null)
            {
                this._stringCache[count] = valueStr;
            }
            this._indexLookup[valueStr] = count;
            this._changed = true;
            return (int) count;
        }

        public void Save()
        {
            if (this._changed)
            {
                if (this._ssPart != null)
                {
                    string idOfPart = this._wpart.GetIdOfPart(this._ssPart);
                    this._wpart.DeletePart(idOfPart);
                }
                using (OpenXmlWriter writer = OpenXmlWriter.Create(this._wpart.AddNewPart<SharedStringTablePart>()))
                {
                    writer.WriteStartElement(new SharedStringTable());
                    if (this._stringCache == null)
                    {
                        string[] strArray = new string[this._indexLookup.Count];
                        foreach (KeyValuePair<string, uint> pair in this._indexLookup)
                        {
                            strArray[(int) ((IntPtr) pair.Value)] = pair.Key;
                        }
                        for (uint i = 0; i < strArray.Length; i++)
                        {
                            writer.WriteStartElement(new SharedStringItem());
                            writer.WriteElement(new Text(strArray[i]));
                            writer.WriteEndElement();
                        }
                    }
                    else
                    {
                        foreach (KeyValuePair<uint, string> pair2 in this._stringCache)
                        {
                            writer.WriteStartElement(new SharedStringItem());
                            writer.WriteElement(new Text(pair2.Value));
                            writer.WriteEndElement();
                        }
                    }
                    writer.WriteEndElement();
                }
            }
        }

        private void StringCacheLazyInit()
        {
            if (this._stringCache == null)
            {
                this._stringCache = new SortedList<uint, string>();
                foreach (KeyValuePair<string, uint> pair in this._indexLookup)
                {
                    this._stringCache[pair.Value] = pair.Key;
                }
            }
        }
    }
}

