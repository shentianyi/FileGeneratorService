namespace OpenExcel.OfficeOpenXml.Internal
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    internal class CellFormulaProxy
    {
        private bool? _AlwaysCalculateArray;
        private bool? _Bx;
        private bool? _CalculateCell;
        private bool? _DataTable2D;
        private bool? _DataTableRow;
        private CellFormulaValues? _FormulaType;
        private bool? _Input1Deleted;
        private bool? _Input2Deleted;
        private string _R1;
        private string _R2;
        private string _Reference;
        private uint? _SharedIndex;
        private string _Text;
        private WorksheetCache _wscache;

        public CellFormulaProxy(WorksheetCache wscache)
        {
            this._wscache = wscache;
        }

        public bool? AlwaysCalculateArray
        {
            get
            {
                return this._AlwaysCalculateArray;
            }
            set
            {
                this._AlwaysCalculateArray = value;
                this._wscache.Modified = true;
            }
        }

        public bool? Bx
        {
            get
            {
                return this._Bx;
            }
            set
            {
                this._Bx = value;
                this._wscache.Modified = true;
            }
        }

        public bool? CalculateCell
        {
            get
            {
                return this._CalculateCell;
            }
            set
            {
                this._CalculateCell = value;
                this._wscache.Modified = true;
            }
        }

        public bool? DataTable2D
        {
            get
            {
                return this._DataTable2D;
            }
            set
            {
                this._DataTable2D = value;
                this._wscache.Modified = true;
            }
        }

        public bool? DataTableRow
        {
            get
            {
                return this._DataTableRow;
            }
            set
            {
                this._DataTableRow = value;
                this._wscache.Modified = true;
            }
        }

        public CellFormulaValues? FormulaType
        {
            get
            {
                return this._FormulaType;
            }
            set
            {
                this._FormulaType = value;
                this._wscache.Modified = true;
            }
        }

        public bool? Input1Deleted
        {
            get
            {
                return this._Input1Deleted;
            }
            set
            {
                this._Input1Deleted = value;
                this._wscache.Modified = true;
            }
        }

        public bool? Input2Deleted
        {
            get
            {
                return this._Input2Deleted;
            }
            set
            {
                this._Input2Deleted = value;
                this._wscache.Modified = true;
            }
        }

        public string R1
        {
            get
            {
                return this._R1;
            }
            set
            {
                this._R1 = value;
                this._wscache.Modified = true;
            }
        }

        public string R2
        {
            get
            {
                return this._R2;
            }
            set
            {
                this._R2 = value;
                this._wscache.Modified = true;
            }
        }

        public string Reference
        {
            get
            {
                return this._Reference;
            }
            set
            {
                this._Reference = value;
                this._wscache.Modified = true;
            }
        }

        public uint? SharedIndex
        {
            get
            {
                return this._SharedIndex;
            }
            set
            {
                this._SharedIndex = value;
                this._wscache.Modified = true;
            }
        }

        public string Text
        {
            get
            {
                return this._Text;
            }
            set
            {
                this._Text = value;
            }
        }
    }
}

