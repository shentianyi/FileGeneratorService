namespace OpenExcel.OfficeOpenXml
{
    using OpenExcel.Common;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;

    public class ExcelCellFormula
    {
        private uint _col;
        private uint _row;
        private ExcelWorksheet _wsheet;

        internal ExcelCellFormula(uint row, uint col, ExcelWorksheet w)
        {
            this._row = row;
            this._col = col;
            this._wsheet = w;
        }

        public void CopyTo(string address)
        {
            RowColumn column = ExcelAddress.ToRowColumn(address);
            this.CopyTo(column.Row, column.Column);
        }

        public void CopyTo(uint targetRow, uint targetCol)
        {
            string str = ExcelFormula.Translate(this.Text, (int) (targetRow - this._row), (int) (targetCol - this._col));
            this._wsheet.Cells[targetRow, targetCol].Formula.Text = str;
        }

        public void Remove()
        {
            CellProxy cell = this._wsheet.GetCell(this._row, this._col);
            if (cell != null)
            {
                cell.RemoveFormula();
            }
        }

        public string Text
        {
            get
            {
                CellProxy cell = this._wsheet.GetCell(this._row, this._col);
                if ((cell != null) && (cell.Formula != null))
                {
                    return cell.Formula.Text;
                }
                return null;
            }
            set
            {
                CellProxy proxy = this._wsheet.EnsureCell(this._row, this._col);
                if (proxy.Formula == null)
                {
                    proxy.CreateFormula();
                    proxy.Formula.Text = value;
                }
            }
        }
    }
}

