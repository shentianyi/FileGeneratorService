namespace OpenExcel.OfficeOpenXml.Internal
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    internal class CellProxy
    {
        private CellValues? _DataType;
        private CellFormulaProxy _Formula;
        private bool? _ShowPhonetic;
        private uint? _StyleIndex;
        private object _Value;
        private uint? _ValueMetaIndex;
        private WorksheetCache _wscache;

        public CellProxy(WorksheetCache wscache)
        { 
            this._wscache = wscache;
        }

        public void CreateFormula()
        {
            this._Formula = new CellFormulaProxy(this._wscache);
        }

        public void RemoveFormula()
        {
            this._Formula = null;
        }

        public CellValues? DataType
        {
            get
            {
                return this._DataType;
            }
            set
            {
                this._DataType = value;
                this._wscache.Modified = true;
            }
        }

        public CellFormulaProxy Formula
        {
            get
            {
                return this._Formula;
            }
        }

        public string SerializedValue
        {
            get
            {
                if (this._Value == null)
                {
                    return "";
                }
                DateTime? nullable = this._Value as DateTime?;
                if (nullable.HasValue)
                {
                    return nullable.Value.ToOADate().ToString();
                }
                return this._Value.ToString();
            }
        }

        public bool? ShowPhonetic
        {
            get
            {
                return this._ShowPhonetic;
            }
            set
            {
                this._ShowPhonetic = value;
                this._wscache.Modified = true;
            }
        }

        public uint? StyleIndex
        {
            get
            {
                return this._StyleIndex;
            }
            set
            {
                this._StyleIndex = value;
                this._wscache.Modified = true;
            }
        }

        public object Value
        {
            get
            {
                return this._Value;
            }
            set
            {
                this._Value = value;
                this._wscache.Modified = true;
            }
        }

        public uint? ValueMetaIndex
        {
            get
            {
                return this._ValueMetaIndex;
            }
            set
            {
                this._ValueMetaIndex = value;
                this._wscache.Modified = true;
            }
        }
    }
}

