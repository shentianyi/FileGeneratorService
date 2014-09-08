namespace OpenExcel.OfficeOpenXml.Internal
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.Common;
    using OpenExcel.OfficeOpenXml;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Threading;

    internal class WorksheetCache
    {
        private SortedList<uint, SortedList<uint, CellProxy>> _cachedCells = new SortedList<uint, SortedList<uint, CellProxy>>();
        private List<OpenXmlElement> _cachedElements = new List<OpenXmlElement>();
        private SortedList<uint, Row> _cachedRows = new SortedList<uint, Row>();
        private ExcelWorksheet _wsheet;

        public WorksheetCache(ExcelWorksheet wsheet)
        {
            this._wsheet = wsheet;
        }

        private void AddCellToCache(CellProxy c, uint rowIdx, uint colIdx)
        {
            SortedList<uint, CellProxy> list;
            this.Modified = true;
            if (!this._cachedCells.TryGetValue(rowIdx, out list))
            {
                list = new SortedList<uint, CellProxy>();
                this._cachedCells[rowIdx] = list;
            }
            list[colIdx] = c;
        }

        private void AddCellToCache(CellProxy c, SortedList<uint, CellProxy> rowCache, uint colIdx)
        {
            this.Modified = true;
            rowCache[colIdx] = c;
        }

        public static RowColumn CacheIndexToRowCol(ulong cacheIdx)
        {
            return new RowColumn { Row = ((uint) (cacheIdx / ((ulong) ExcelConstraints.MaxColumns))) + 1, Column = ((uint) (cacheIdx % ((ulong) ExcelConstraints.MaxColumns))) + 1 };
        }

        public static void CacheIndexToRowCol(ulong cacheIdx, out uint rowIdx, out uint colIdx)
        {
            rowIdx = ((uint) (cacheIdx / ((ulong) ExcelConstraints.MaxColumns))) + 1;
            colIdx = ((uint) (cacheIdx % ((ulong) ExcelConstraints.MaxColumns))) + 1;
        }

        private IEnumerable<KeyValuePair<ulong, CellProxy>> Cells()
        {
            foreach (KeyValuePair<uint, SortedList<uint, CellProxy>> iteratorVariable0 in this._cachedCells)
            {
                foreach (KeyValuePair<uint, CellProxy> iteratorVariable1 in iteratorVariable0.Value)
                {
                    ulong key = RowColToCacheIndex(iteratorVariable0.Key, iteratorVariable1.Key);
                    yield return new KeyValuePair<ulong, CellProxy>(key, iteratorVariable1.Value);
                }
            }
        }

        private void CreateColumnCopies(uint colFromIdx, int numOfColumns, Func<CellProxy, CellProxy> fnCreate)
        {
            Action<uint, uint> action = delegate (uint rowIdx, uint colIdx) {
                RowColumn column = new RowColumn {
                    Row = rowIdx,
                    Column = colIdx
                };
                if (column.Column == colFromIdx)
                {
                    foreach (int num in Enumerable.Range(((int) colFromIdx) + 1, numOfColumns))
                    {
                        RowColumn column2 = new RowColumn {
                            Row = column.Row,
                            Column = (uint) num
                        };
                        CellProxy cell = this.GetCell(column.Row, column.Column);
                        if (cell != null)
                        {
                            CellProxy c = fnCreate(cell);
                            if (c != null)
                            {
                                this.AddCellToCache(c, column2.Row, column2.Column);
                            }
                        }
                    }
                }
            };
            this.LoopCells(action);
        }

        private void CreateRowCopies(uint rowFromIdx, int numOfRows, Func<CellProxy, CellProxy> fnCreate)
        {
            SortedList<uint, CellProxy> list;
            if (this._cachedCells.TryGetValue(rowFromIdx, out list))
            {
                Row row;
                this._cachedRows.TryGetValue(rowFromIdx, out row);
                foreach (int num in Enumerable.Range(((int) rowFromIdx) + 1, numOfRows))
                {
                    if (row != null)
                    {
                        Row row2 = (Row) row.CloneNode(false);
                        row2.RowIndex = (UInt32Value) (uint)num;
                        this._cachedRows.Add((uint) row2.RowIndex, row2);
                    }
                    foreach (uint num2 in list.Keys)
                    {
                        RowColumn column = new RowColumn {
                            Row = rowFromIdx,
                            Column = num2
                        };
                        RowColumn column2 = new RowColumn {
                            Row = (uint) num,
                            Column = num2
                        };
                        CellProxy cell = this.GetCell(column.Row, column.Column);
                        CellProxy c = fnCreate(cell);
                        if (c != null)
                        {
                            this.AddCellToCache(c, column2.Row, column2.Column);
                        }
                    }
                }
            }
        }

        public void DeleteSingleSpanColumn(uint col)
        {
            Func<Column, bool> func = null;
            Columns firstElement = this.GetFirstElement<Columns>();
            if (firstElement != null)
            {
                if (func == null)
                {
                    func = c => (c.Min == col) && (c.Max == col);
                }
                Column column = Enumerable.Where<Column>(firstElement.Elements<Column>(), func).FirstOrDefault<Column>();
                if (column != null)
                {
                    column.Remove();
                }
            }
        }

        public CellProxy EnsureCell(uint row, uint col)
        {
            CellProxy c = null;
            SortedList<uint, CellProxy> list;
            if (!this._cachedCells.TryGetValue(row, out list))
            {
                c = new CellProxy(this);
                this.AddCellToCache(c, row, col);
                return c;
            }
            if (!list.TryGetValue(col, out c))
            {
                c = new CellProxy(this);
                if (list != null)
                {
                    this.AddCellToCache(c, list, col);
                }
            }
            return c;
        }

        public Row EnsureRow(uint row)
        {
            Row row2 = null;
            if (!this._cachedRows.TryGetValue(row, out row2))
            {
                Row row3 = new Row {
                    RowIndex = row
                };
                this._cachedRows[row] = row2 = row3;
            }
            return row2;
        }

        public Column EnsureSingleSpanColumn(uint col)
        {
            Func<Column, bool> func = null;
            Columns firstElement = this.GetFirstElement<Columns>();
            if (firstElement == null)
            {
                firstElement = new Columns();
                this._cachedElements.Add(firstElement);
            }
            Column refChild = (from c in firstElement.Elements<Column>()
                where (c.Min <= col) && (c.Max >= col)
                select c).FirstOrDefault<Column>();
            if (refChild != null)
            {
                if (refChild.Min < col)
                {
                    Column column2 = (Column) refChild.CloneNode(false);
                    column2.Min = refChild.Min;
                    column2.Max = col - 1;
                    refChild.Min = col;
                    firstElement.InsertBefore<Column>(column2, refChild);
                }
                if (refChild.Max > col)
                {
                    Column column3 = (Column) refChild.CloneNode(false);
                    column3.Min = col + 1;
                    column3.Max = refChild.Max;
                    refChild.Max = col;
                    firstElement.InsertAfter<Column>(column3, refChild);
                }
                return refChild;
            }
            Column newChild = new Column {
                Min = col,
                Max = col
            };
            if (func == null)
            {
                func = c => c.Min > col;
            }
            Column column5 = Enumerable.Where<Column>(firstElement.Elements<Column>(), func).FirstOrDefault<Column>();
            if (column5 != null)
            {
                firstElement.InsertBefore<Column>(newChild, column5);
                return newChild;
            }
            firstElement.Append(new OpenXmlElement[] { newChild });
            return newChild;
        }

        private IEnumerable<OpenXmlAttribute> EnumCellProxyAttributes(uint row, uint col, CellProxy cellProxy)
        {
            OpenXmlAttribute iteratorVariable0 = new OpenXmlAttribute {
                LocalName = "r",
                Value = RowColumn.ToAddress(row, col)
            };
            yield return iteratorVariable0;
            if (cellProxy.DataType.HasValue)
            {
                OpenXmlAttribute iteratorVariable1 = new OpenXmlAttribute {
                    LocalName = "t",
                    Value = this.STCellType(cellProxy.DataType.Value)
                };
                yield return iteratorVariable1;
            }
            if (cellProxy.StyleIndex.HasValue)
            {
                OpenXmlAttribute iteratorVariable2 = new OpenXmlAttribute {
                    LocalName = "s",
                    Value = cellProxy.StyleIndex.Value.ToString()
                };
                yield return iteratorVariable2;
            }
        }

        private IEnumerable<OpenXmlAttribute> EnumRowAttributes(uint thisRowIdx, SortedList<uint, CellProxy> cells)
        {
            Row iteratorVariable0;
            bool iteratorVariable1 = false;
            if (this._cachedRows.TryGetValue(thisRowIdx, out iteratorVariable0))
            {
                IEnumerable<OpenXmlAttribute> iteratorVariable2 = from a in iteratorVariable0.GetAttributes()
                    where (a.LocalName != "r") && (a.LocalName != "spans")
                    select a;
                foreach (OpenXmlAttribute iteratorVariable3 in iteratorVariable2)
                {
                    yield return iteratorVariable3;
                    iteratorVariable1 = true;
                }
            }
            if (iteratorVariable1 || (cells.Count != 0))
            {
                OpenXmlAttribute iteratorVariable7 = new OpenXmlAttribute {
                    LocalName = "r",
                    Value = thisRowIdx.ToString()
                };
                yield return iteratorVariable7;
                if (cells.Count > 0)
                {
                    uint iteratorVariable4 = cells.Keys.First<uint>();
                    uint iteratorVariable5 = cells.Keys.Last<uint>();
                    OpenXmlAttribute iteratorVariable6 = new OpenXmlAttribute {
                        LocalName = "spans",
                        Value = iteratorVariable4 + ":" + iteratorVariable5
                    };
                    yield return iteratorVariable6;
                }
            }
            else
            {
                yield break;
            }
        }

        public CellProxy GetCell(uint row, uint col)
        {
            CellProxy proxy;
            SortedList<uint, CellProxy> list;
            if (this._cachedCells.TryGetValue(row, out list) && list.TryGetValue(col, out proxy))
            {
                return proxy;
            }
            return null;
        }

        public Column GetContainingColumn(uint col)
        {
            Func<Column, bool> func = null;
            Columns firstElement = this.GetFirstElement<Columns>();
            if (firstElement == null)
            {
                return null;
            }
            if (func == null)
            {
                func = c => (c.Min <= col) && (c.Max >= col);
            }
            return Enumerable.Where<Column>(firstElement.Elements<Column>(), func).FirstOrDefault<Column>();
        }

        private IEnumerable<T> GetElements<T>() where T: OpenXmlElement
        {
            return (from e in this._cachedElements
                where e.GetType() == typeof(T)
                select (T) e);
        }

        private IEnumerable<OpenXmlElement> GetElementsByTagName(string name)
        {
            return (from e in this._cachedElements
                where e.LocalName == name
                select e);
        }

        private T GetFirstElement<T>() where T: OpenXmlElement
        {
            return (T)(from e in this._cachedElements
                where e.GetType() == typeof(T)
                select e).FirstOrDefault<OpenXmlElement>();
        }

        public Row GetRow(uint row)
        {
            Row row2;
            if (this._cachedRows.TryGetValue(row, out row2))
            {
                return row2;
            }
            return null;
        }

        public void InsertOrDeleteColumns(uint colStart, int colDelta, bool copyPreviousStyle)
        {
            Func<Column, bool> func = null;
            Func<CellProxy, CellProxy> fnCreate = null;
            if (colDelta != 0)
            {
                this.Modified = true;
                Action<uint, uint> action = delegate (uint rowIdx, uint colIdx) {
                    if (colIdx >= colStart)
                    {
                        CellProxy c = this.RemoveCellFromCache(rowIdx, colIdx);
                        int num = ((int) colIdx) + colDelta;
                        if ((num >= colStart) && (num >= 1))
                        {
                            if (colIdx >= ExcelConstraints.MaxColumns)
                            {
                                throw new InvalidOperationException("Max number of columns exceeded");
                            }
                            this.AddCellToCache(c, rowIdx, (uint) num);
                        }
                    }
                };
                if (colDelta > 0)
                {
                    this.LoopCellsReverse(action);
                }
                else
                {
                    this.LoopCells(action);
                }
                Columns firstElement = this.GetFirstElement<Columns>();
                if (firstElement != null)
                {
                    List<Column> list = new List<Column>();
                    foreach (Column column in firstElement)
                    {
                        if (column.Min >= colStart)
                        {
                            column.Min = (uint) Math.Max((long) 0L, (long) (((long) ((ulong) column.Min)) + colDelta));
                        }
                        if (column.Max >= colStart)
                        {
                            column.Max = (uint) Math.Max((long) 0L, (long) (((long) ((ulong) column.Max)) + colDelta));
                        }
                        if ((column.Min <= 0) || (column.Max < column.Min))
                        {
                            list.Add(column);
                        }
                    }
                    foreach (Column column2 in list)
                    {
                        column2.Remove();
                    }
                    if (func == null)
                    {
                        func = col => col.Max == (colStart - 1);
                    }
                    Column column3 = Enumerable.Where<Column>(firstElement.Elements<Column>(), func).FirstOrDefault<Column>();
                    if (column3 != null)
                    {
                        column3.Max = (uint) (((ulong) column3.Max) + (ulong)colDelta);
                    }
                    if (firstElement.ChildElements.Count == 0)
                    {
                        this._cachedElements.Remove(firstElement);
                    }
                }
                if (colDelta > 0)
                {
                    if (copyPreviousStyle)
                    {
                        if (fnCreate == null)
                        {
                            fnCreate = delegate (CellProxy cOld) {
                                if (cOld.StyleIndex.HasValue)
                                {
                                    return new CellProxy(this) { StyleIndex = cOld.StyleIndex };
                                }
                                return null;
                            };
                        }
                        this.CreateColumnCopies(colStart - 1, colDelta, fnCreate);
                    }
                    else
                    {
                        for (uint i = colStart; i < (colStart + colDelta); i++)
                        {
                            this.EnsureSingleSpanColumn(i);
                            this.DeleteSingleSpanColumn(i);
                        }
                    }
                }
            }
        }

        public void InsertOrDeleteRows(uint rowStart, int rowDelta, bool copyPreviousStyle)
        {
            Func<CellProxy, CellProxy> fnCreate = null;
            if (rowDelta != 0)
            {
                IList<uint> list;
                IEnumerable<uint> enumerable;
                this.Modified = true;
                if (rowDelta > 0)
                {
                    list = this._cachedCells.Keys.Reverse<uint>().ToList<uint>();
                }
                else
                {
                    list = this._cachedCells.Keys.ToList<uint>();
                }
                SortedList<uint, SortedList<uint, CellProxy>> list2 = new SortedList<uint, SortedList<uint, CellProxy>>();
                foreach (uint num in list)
                {
                    uint num2;
                    if (num >= rowStart)
                    {
                        num2 = num + ((uint) rowDelta);
                    }
                    else
                    {
                        num2 = num;
                    }
                    list2[num2] = this._cachedCells[num];
                }
                this._cachedCells = list2;
                if (rowDelta > 0)
                {
                    enumerable = (from k in this._cachedRows.Keys
                        where k >= rowStart
                        select k).Reverse<uint>().ToList<uint>();
                }
                else
                {
                    enumerable = (from k in this._cachedRows.Keys
                        where k >= rowStart
                        select k).ToList<uint>();
                }
                foreach (uint num3 in enumerable)
                {
                    Row row = this._cachedRows[num3];
                    int num4 = (int)(row.RowIndex.Value + rowDelta);
                    this._cachedRows.Remove(num3);
                    if ((num4 >= rowStart) && (num4 >= 1))
                    {
                        row.RowIndex = (UInt32Value)(uint) num4;
                        this._cachedRows[(uint) row.RowIndex] = row;
                    }
                }
                if ((rowDelta > 0) && copyPreviousStyle)
                {
                    if (fnCreate == null)
                    {
                        fnCreate = delegate (CellProxy cOld) {
                            if (cOld.StyleIndex.HasValue)
                            {
                                return new CellProxy(this) { StyleIndex = cOld.StyleIndex };
                            }
                            return null;
                        };
                    }
                    this.CreateRowCopies(rowStart - 1, rowDelta, fnCreate);
                }
            }
        }

        public void Load()
        {
            if (this._wsheet.GetOWorksheetPart() != null)
            {
                Action<Cell> action = delegate (Cell cell) {
                    RowColumn column = ExcelAddress.ToRowColumn((string) cell.CellReference);
                    CellProxy proxy = this.EnsureCell(column.Row, column.Column);
                    if (cell.DataType != null)
                    {
                        proxy.DataType = new CellValues?(cell.DataType.Value);
                        if (((CellValues) cell.DataType.Value) == CellValues.InlineString)
                        {
                            proxy.Value = cell.InlineString.Text.Text;
                        }
                        else if (cell.CellValue != null)
                        {
                            proxy.Value = cell.CellValue.Text;
                        }
                        else
                        {
                            proxy.Value = string.Empty;
                        }
                    }
                    else if (cell.CellValue != null)
                    {
                        proxy.Value = cell.CellValue.Text;
                    }
                    if (cell.StyleIndex != null)
                    {
                        proxy.StyleIndex = new uint?((uint) cell.StyleIndex);
                    }
                    if (cell.ShowPhonetic != null)
                    {
                        proxy.ShowPhonetic = new bool?((bool) cell.ShowPhonetic);
                    }
                    if (cell.ValueMetaIndex != null)
                    {
                        proxy.ValueMetaIndex = new uint?((uint) cell.ValueMetaIndex);
                    }
                    if (cell.CellFormula != null)
                    {
                        proxy.CreateFormula();
                        proxy.Formula.Text = cell.CellFormula.Text;
                        proxy.Formula.R1 = (string) cell.CellFormula.R1;
                        proxy.Formula.R2 = (string) cell.CellFormula.R2;
                        proxy.Formula.Reference = (string) cell.CellFormula.Reference;
                        if (cell.CellFormula.AlwaysCalculateArray != null)
                        {
                            proxy.Formula.AlwaysCalculateArray = new bool?((bool) cell.CellFormula.AlwaysCalculateArray);
                        }
                        if (cell.CellFormula.Bx != null)
                        {
                            proxy.Formula.Bx = new bool?((bool) cell.CellFormula.Bx);
                        }
                        if (cell.CellFormula.CalculateCell != null)
                        {
                            proxy.Formula.CalculateCell = new bool?((bool) cell.CellFormula.CalculateCell);
                        }
                        if (cell.CellFormula.DataTable2D != null)
                        {
                            proxy.Formula.DataTable2D = new bool?((bool) cell.CellFormula.DataTable2D);
                        }
                        if (cell.CellFormula.DataTableRow != null)
                        {
                            proxy.Formula.DataTableRow = new bool?((bool) cell.CellFormula.DataTableRow);
                        }
                        if (cell.CellFormula.FormulaType != null)
                        {
                            proxy.Formula.FormulaType = new CellFormulaValues?((CellFormulaValues) cell.CellFormula.FormulaType);
                        }
                        if (cell.CellFormula.Input1Deleted != null)
                        {
                            proxy.Formula.Input1Deleted = new bool?((bool) cell.CellFormula.Input1Deleted);
                        }
                        if (cell.CellFormula.Input2Deleted != null)
                        {
                            proxy.Formula.Input2Deleted = new bool?((bool) cell.CellFormula.Input2Deleted);
                        }
                        if (cell.CellFormula.SharedIndex != null)
                        {
                            proxy.Formula.SharedIndex = new uint?((uint) cell.CellFormula.SharedIndex);
                        }
                        proxy.Value = null;
                    }
                };
                OpenXmlReader reader = OpenXmlReader.Create(this._wsheet.GetOWorksheetPart());
                bool flag = false;
                while (reader.Read())
                {
                    if (flag && reader.IsStartElement)
                    {
                        if ((reader.ElementType == typeof(Row)) && reader.IsStartElement)
                        {
                            Row row = (Row) reader.LoadCurrentElement();
                            if ((from a in row.GetAttributes()
                                let ln = a.LocalName
                                where (ln != "r") && (ln != "spans")
                                select a).Any<OpenXmlAttribute>())
                            {
                                this._cachedRows.Add((uint) row.RowIndex, (Row) row.CloneNode(false));
                            }
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                action(cell);
                            }
                        }
                        else if (reader.ElementType == typeof(SheetData))
                        {
                            bool isStartElement = reader.IsStartElement;
                        }
                        else if (reader.IsStartElement)
                        {
                            OpenXmlElement item = reader.LoadCurrentElement();
                            this._cachedElements.Add(item);
                        }
                    }
                    else if (reader.ElementType == typeof(Worksheet))
                    {
                        flag = reader.IsStartElement;
                    }
                }
                this.Modified = false;
            }
        }

        private void LoopCells(Action<uint, uint> action)
        {
            IList<uint> source = this._cachedCells.Keys.ToList<uint>();
            uint num = source.First<uint>();
            uint num2 = source.Last<uint>();
            for (uint i = num; i <= num2; i++)
            {
                SortedList<uint, CellProxy> list2;
                if (this._cachedCells.TryGetValue(i, out list2))
                {
                    IList<uint> list3 = list2.Keys.ToList<uint>();
                    uint num4 = list3.First<uint>();
                    uint num5 = list3.Last<uint>();
                    for (uint j = num4; j <= num5; j++)
                    {
                        if (list2.ContainsKey(j))
                        {
                            action(i, j);
                        }
                    }
                }
            }
        }

        private void LoopCellsReverse(Action<uint, uint> action)
        {
            IList<uint> source = this._cachedCells.Keys.ToList<uint>();
            uint num = source.First<uint>();
            for (uint i = source.Last<uint>(); i >= num; i--)
            {
                SortedList<uint, CellProxy> list2;
                if (this._cachedCells.TryGetValue(i, out list2))
                {
                    IList<uint> list3 = list2.Keys.ToList<uint>();
                    uint num4 = list3.First<uint>();
                    for (uint j = list3.Last<uint>(); j >= num4; j--)
                    {
                        if (list2.ContainsKey(j))
                        {
                            action(i, j);
                        }
                    }
                }
            }
        }

        public void RecalcCellReferences(SheetChange sheetChange)
        {
            foreach (KeyValuePair<ulong, CellProxy> pair in this.Cells())
            {
                CellProxy proxy = pair.Value;
                if (proxy.Formula != null)
                {
                    proxy.Formula.Text = ExcelFormula.TranslateForSheetChange(proxy.Formula.Text, sheetChange, this._wsheet.Name);
                    proxy.Formula.R1 = ExcelRange.TranslateForSheetChange(proxy.Formula.R1, sheetChange, this._wsheet.Name);
                    proxy.Formula.R2 = ExcelRange.TranslateForSheetChange(proxy.Formula.R2, sheetChange, this._wsheet.Name);
                    proxy.Formula.Reference = ExcelRange.TranslateForSheetChange(proxy.Formula.Reference, sheetChange, this._wsheet.Name);
                }
            }
            List<ConditionalFormatting> list = new List<ConditionalFormatting>();
            foreach (ConditionalFormatting formatting in this.GetElements<ConditionalFormatting>())
            {
                bool flag = false;
                List<StringValue> list2 = new List<StringValue>();
                foreach (StringValue value2 in formatting.SequenceOfReferences.Items)
                {
                    string str = ExcelRange.TranslateForSheetChange(value2.Value, sheetChange, this._wsheet.Name);
                    if (!str.StartsWith("#"))
                    {
                        list2.Add(new StringValue(str));
                    }
                    else
                    {
                        list.Add(formatting);
                        flag = true;
                        break;
                    }
                }
                if (flag)
                {
                    break;
                }
                formatting.SequenceOfReferences = new ListValue<StringValue>(list2);
                foreach (Formula formula in formatting.Descendants<Formula>())
                {
                    formula.Text = ExcelFormula.TranslateForSheetChange(formula.Text, sheetChange, this._wsheet.Name);
                }
            }
            foreach (ConditionalFormatting formatting2 in list)
            {
                this.RemoveElement(formatting2);
            }
        }

        private CellProxy RemoveCellFromCache(uint rowIdx, uint colIdx)
        {
            SortedList<uint, CellProxy> list;
            this.Modified = true;
            CellProxy proxy = null;
            if (this._cachedCells.TryGetValue(rowIdx, out list) && list.TryGetValue(colIdx, out proxy))
            {
                list.Remove(colIdx);
                return proxy;
            }
            return null;
        }

        private void RemoveElement(OpenXmlElement e)
        {
            this.Modified = true;
            this._cachedElements.Remove(e);
        }

        public static ulong RowColToCacheIndex(uint rowIdx, uint colIdx)
        {
            return (ulong) (((rowIdx - 1) * ExcelConstraints.MaxColumns) + (colIdx - 1));
        }

        private string STCellType(CellValues v)
        {
            switch (v)
            {
                case CellValues.Boolean:
                    return "b";

                case CellValues.Number:
                    return "n";

                case CellValues.Error:
                    return "e";

                case CellValues.SharedString:
                    return "s";

                case CellValues.String:
                    return "str";

                case CellValues.InlineString:
                    return "inlineStr";

                case CellValues.Date:
                    return "d";
            }
            return "";
        }

        private void WriteSheetData(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new SheetData());
            foreach (KeyValuePair<uint, SortedList<uint, CellProxy>> pair in this._cachedCells)
            {
                uint key = pair.Key;
                SortedList<uint, CellProxy> cells = pair.Value;
                List<OpenXmlAttribute> attributes = this.EnumRowAttributes(key, cells).ToList<OpenXmlAttribute>();
                if ((attributes.Count > 0) || (cells.Count > 0))
                {
                    writer.WriteStartElement(new Row(), attributes);
                    foreach (KeyValuePair<uint, CellProxy> pair2 in cells)
                    {
                        uint col = pair2.Key;
                        CellProxy cellProxy = pair2.Value;
                        writer.WriteStartElement(new Cell(), this.EnumCellProxyAttributes(key, col, cellProxy));
                        if (cellProxy.Formula != null)
                        {
                            CellFormula elementObject = new CellFormula(cellProxy.Formula.Text);
                            if (cellProxy.Formula.R1 != null)
                            {
                                elementObject.R1 = cellProxy.Formula.R1;
                            }
                            if (cellProxy.Formula.R2 != null)
                            {
                                elementObject.R2 = cellProxy.Formula.R2;
                            }
                            if (cellProxy.Formula.Reference != null)
                            {
                                elementObject.Reference = cellProxy.Formula.Reference;
                            }
                            if (cellProxy.Formula.AlwaysCalculateArray.HasValue)
                            {
                                bool? alwaysCalculateArray = cellProxy.Formula.AlwaysCalculateArray;
                                elementObject.AlwaysCalculateArray = alwaysCalculateArray.HasValue ? ((BooleanValue) alwaysCalculateArray.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.Bx.HasValue)
                            {
                                bool? bx = cellProxy.Formula.Bx;
                                elementObject.Bx = bx.HasValue ? ((BooleanValue) bx.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.CalculateCell.HasValue)
                            {
                                bool? calculateCell = cellProxy.Formula.CalculateCell;
                                elementObject.CalculateCell = calculateCell.HasValue ? ((BooleanValue) calculateCell.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.DataTable2D.HasValue)
                            {
                                bool? nullable8 = cellProxy.Formula.DataTable2D;
                                elementObject.DataTable2D = nullable8.HasValue ? ((BooleanValue) nullable8.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.DataTableRow.HasValue)
                            {
                                bool? dataTableRow = cellProxy.Formula.DataTableRow;
                                elementObject.DataTableRow = dataTableRow.HasValue ? ((BooleanValue) dataTableRow.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.FormulaType.HasValue)
                            {
                                CellFormulaValues? formulaType = cellProxy.Formula.FormulaType;
                                elementObject.FormulaType = formulaType.HasValue ? ((EnumValue<CellFormulaValues>) formulaType.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.Input1Deleted.HasValue)
                            {
                                bool? nullable14 = cellProxy.Formula.Input1Deleted;
                                elementObject.Input1Deleted = nullable14.HasValue ? ((BooleanValue) nullable14.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.Input2Deleted.HasValue)
                            {
                                bool? nullable16 = cellProxy.Formula.Input2Deleted;
                                elementObject.Input2Deleted = nullable16.HasValue ? ((BooleanValue) nullable16.GetValueOrDefault()) : null;
                            }
                            if (cellProxy.Formula.SharedIndex.HasValue)
                            {
                                uint? sharedIndex = cellProxy.Formula.SharedIndex;
                                elementObject.SharedIndex = sharedIndex.HasValue ? ((UInt32Value) sharedIndex.GetValueOrDefault()) : null;
                            }
                            writer.WriteElement(elementObject);
                        }
                        if (cellProxy.Value != null)
                        {
                            writer.WriteElement(new CellValue(cellProxy.SerializedValue));
                        }
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
            }
            writer.WriteEndElement();
        }

        public void WriteWorksheetPart(OpenXmlWriter writer)
        {
            foreach (KeyValuePair<uint, SortedList<uint, CellProxy>> pair in this._cachedCells.ToList<KeyValuePair<uint, SortedList<uint, CellProxy>>>())
            {
                if (pair.Value.Count == 0)
                {
                    this._cachedCells.Remove(pair.Key);
                }
            }
            foreach (uint num in this._cachedRows.Keys)
            {
                if (!this._cachedCells.ContainsKey(num))
                {
                    this._cachedCells[num] = new SortedList<uint, CellProxy>();
                }
            }
            uint maxValue = uint.MaxValue;
            uint num3 = uint.MaxValue;
            uint row = 0;
            uint num5 = 0;
            foreach (KeyValuePair<uint, SortedList<uint, CellProxy>> pair2 in this._cachedCells)
            {
                uint key = pair2.Key;
                SortedList<uint, CellProxy> list = pair2.Value;
                if (maxValue == uint.MaxValue)
                {
                    maxValue = key;
                }
                row = key;
                if (list.Count > 0)
                {
                    num3 = Math.Min(num3, list.Keys.First<uint>());
                    num5 = Math.Max(num5, list.Keys.Last<uint>());
                }
            }
            string str = null;
            string str2 = null;
            if ((maxValue < uint.MaxValue) && (num3 < uint.MaxValue))
            {
                str = RowColumn.ToAddress(maxValue, num3);
                if ((maxValue != row) || (num3 != num5))
                {
                    str2 = RowColumn.ToAddress(row, num5);
                }
            }
            else
            {
                str = "A1";
            }
            writer.WriteStartDocument();
            writer.WriteStartElement(new Worksheet());
            foreach (string str3 in SchemaInfo.WorksheetChildSequence)
            {
                switch (str3)
                {
                    case "sheetData":
                        this.WriteSheetData(writer);
                        break;

                    case "dimension":
                    {
                        string str4 = str + ((str2 != null) ? (":" + str2) : "");
                        SheetDimension elementObject = new SheetDimension {
                            Reference = str4
                        };
                        writer.WriteElement(elementObject);
                        break;
                    }
                    case "sheetViews":
                    {
                        SheetViews firstElement = this.GetFirstElement<SheetViews>();
                        if (firstElement != null)
                        {
                            foreach (SheetView view in firstElement.Elements<SheetView>())
                            {
                                foreach (Selection selection in view.Elements<Selection>())
                                {
                                    if (maxValue < uint.MaxValue)
                                    {
                                        selection.ActiveCell = str;
                                        selection.SequenceOfReferences = new ListValue<StringValue>(new StringValue[] { new StringValue(str) });
                                    }
                                    else
                                    {
                                        selection.Remove();
                                    }
                                }
                            }
                            writer.WriteElement(firstElement);
                        }
                        break;
                    }
                    default:
                        foreach (OpenXmlElement element in this.GetElementsByTagName(str3))
                        {
                            writer.WriteElement(element);
                        }
                        break;
                }
            }
            writer.WriteEndElement();
        }

        public bool Modified { get; set; }



    }
}

