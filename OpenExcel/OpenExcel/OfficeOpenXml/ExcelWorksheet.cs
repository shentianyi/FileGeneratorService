namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.Common;
    using OpenExcel.OfficeOpenXml.Internal;
    using System;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Xml;

    public class ExcelWorksheet
    {
        private string _name;
        private WorksheetCache _sheetCache;

        internal ExcelWorksheet(string name, ExcelDocument parentDoc)
        {
            this._name = name;
            this.Document = parentDoc;
            this.Rows = new ExcelRows(this);
            this.Columns = new ExcelColumns(this);
            this.Cells = new ExcelCells(this);
            this._sheetCache = new WorksheetCache(this);
            this._sheetCache.Load();
        }

        public void DeleteColumn(string columnName)
        {
            this.DeleteColumns(ExcelAddress.ColumnNameToIndex(columnName), 1);
        }

        public void DeleteColumn(uint col)
        {
            this.DeleteColumns(col, 1);
        }

        internal void DeleteColumnDefinition(uint col)
        {
            this._sheetCache.DeleteSingleSpanColumn(col);
        }

        public void DeleteColumns(string columnName, int qty)
        {
            this.DeleteColumns(ExcelAddress.ColumnNameToIndex(columnName), qty);
        }

        public void DeleteColumns(uint col, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteColumns(col, -qty, false);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    ColumnStart = col,
                    ColumnDelta = -qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        public void DeleteRow(uint row)
        {
            this.DeleteRows(row, 1);
        }

        public void DeleteRows(uint rowStart, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteRows(rowStart, -qty, false);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    RowStart = rowStart,
                    RowDelta = -qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        internal CellProxy EnsureCell(uint row, uint col)
        {
            return this._sheetCache.EnsureCell(row, col);
        }

        internal Column EnsureColumnDefinition(uint col)
        {
            return this._sheetCache.EnsureSingleSpanColumn(col);
        }

        internal Row EnsureRow(uint row)
        {
            return this._sheetCache.EnsureRow(row);
        }

        internal CellProxy GetCell(uint row, uint col)
        {
            return this._sheetCache.GetCell(row, col);
        }

       
        internal Column GetColumnDefinition(uint col)
        {
            return this._sheetCache.GetContainingColumn(col);
        }

        internal WorksheetPart GetOWorksheetPart()
        {
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            Sheet sheet = (from s in workbookPart.Workbook.Sheets.Elements<Sheet>()
                where s.Name == this.Name
                select s).FirstOrDefault<Sheet>();
            if (sheet != null)
            {
                string id = sheet.Id.Value;
                return (WorksheetPart) workbookPart.GetPartById(id);
            }
            return null;
        }

        internal Row GetRow(uint row)
        {
            return this._sheetCache.GetRow(row);
        }

        internal Row GetOpenRow(uint rowIndex)
        { 
            Worksheet worksheet = this.Document.GetOSpreadsheet().WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
          int i=  worksheet.GetFirstChild<SheetData>().
            Elements<Row>().Count();
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

         // get openxml cell
        internal Cell GetOpenCell(uint rowIndex,uint colIndex)
        {
            Row row = GetRow(rowIndex);
            if (row == null)
                return null;
            return row.Elements<Cell>().ToList()[(int)colIndex];
        }
        // get cellname
        internal string GetCellName(uint rowIndex, uint colIndex)
        {
            Cell cell = GetOpenCell(rowIndex, colIndex);
            return cell == null ? null : cell.CellReference.Value;
        }

        // merge cells
        public void MergeCells(uint row, uint col, uint endRow, uint endCol)
        {
            string cell1Name = GetCellName(row+1, col);
            string cell2Name = GetCellName(endRow+1, endCol);
            MergeTwoCells(cell1Name, cell2Name);
           // Worksheet worksheet = this.Document.GetOSpreadsheet().WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
           // worksheet.Save();
        }

        public void InsertColumn(string columnName)
        {
            this.InsertColumns(ExcelAddress.ColumnNameToIndex(columnName), 1);
        }

        public void InsertColumn(uint col)
        {
            this.InsertColumns(col, 1);
        }

        public void InsertColumns(string columnName, int qty)
        {
            this.InsertColumns(ExcelAddress.ColumnNameToIndex(columnName), qty);
        }

        public void InsertColumns(uint colStart, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteColumns(colStart, qty, true);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    ColumnStart = colStart,
                    ColumnDelta = qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        public void InsertRow(uint row)
        {
            this.InsertRows(row, 1);
        }

        public void InsertRows(uint rowStart, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteRows(rowStart, qty, true);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    RowStart = rowStart,
                    RowDelta = qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        public void MergeTwoCells(string cell1Name, string cell2Name)
        {
            MergeCells cells;
            Worksheet worksheet = this.Document.GetOSpreadsheet().WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
            if (worksheet.Elements<MergeCells>().Count<MergeCells>() > 0)
            {
                cells = worksheet.Elements<MergeCells>().First<MergeCells>();
            }
            else
            {
                cells = new MergeCells();
                if (worksheet.Elements<CustomSheetView>().Count<CustomSheetView>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<CustomSheetView>().First<CustomSheetView>());
                }
                else if (worksheet.Elements<DataConsolidate>().Count<DataConsolidate>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<DataConsolidate>().First<DataConsolidate>());
                }
                else if (worksheet.Elements<SortState>().Count<SortState>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<SortState>().First<SortState>());
                }
                else if (worksheet.Elements<AutoFilter>().Count<AutoFilter>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<AutoFilter>().First<AutoFilter>());
                }
                else if (worksheet.Elements<Scenarios>().Count<Scenarios>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<Scenarios>().First<Scenarios>());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count<ProtectedRanges>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<ProtectedRanges>().First<ProtectedRanges>());
                }
                else if (worksheet.Elements<SheetProtection>().Count<SheetProtection>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<SheetProtection>().First<SheetProtection>());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count<SheetCalculationProperties>() > 0)
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<SheetCalculationProperties>().First<SheetCalculationProperties>());
                }
                else
                {
                    worksheet.InsertAfter<MergeCells>(cells, worksheet.Elements<SheetData>().First<SheetData>());
                }
            }
            MergeCell cell = new MergeCell {
                Reference = new StringValue(cell1Name + ":" + cell2Name)
            };
            cells.Append(new OpenXmlElement[] { cell });
        }

        /// <summary>
        /// merge cell by ws
        /// </summary>
        /// Sets up a region of merged cells.
        public void MergeCells(ExcelCell startCell, int numRows, int numCols)
        {
            //MergeCells(startCell, Cell(startCell.Row + numRows, startCell.Column + numCols));
        }


        public void PushColumn(string columnName)
        {
            this.PushColumns(ExcelAddress.ColumnNameToIndex(columnName), 1);
        }

        public void PushColumns(string columnName, int qty)
        {
            this.PushColumns(ExcelAddress.ColumnNameToIndex(columnName), qty);
        }

        public void PushColumns(uint colStart, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteColumns(colStart, qty, false);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    ColumnStart = colStart,
                    ColumnDelta = qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        public void PushRow(uint row)
        {
            this.PushRows(row, 1);
        }

        public void PushRows(uint rowStart, int qty)
        {
            if (qty < 1)
            {
                throw new ArgumentException("Quantity cannot be less than 0");
            }
            if (qty != 0)
            {
                this._sheetCache.InsertOrDeleteRows(rowStart, qty, false);
                SheetChange sheetChange = new SheetChange {
                    SheetName = this.Name,
                    RowStart = rowStart,
                    RowDelta = qty
                };
                this.Document.RecalcCellReferences(sheetChange);
            }
        }

        internal void RecalcCellReferences(SheetChange sheetChange)
        {
            this._sheetCache.RecalcCellReferences(sheetChange);
        }

        public bool Save()
        {
            if (!this.Modified && !this._sheetCache.Modified)
            {
                return false;
            }
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            Sheet sheet = (from s in workbookPart.Workbook.Sheets.Elements<Sheet>()
                where s.Name == this.Name
                select s).First<Sheet>();
            if (sheet != null)
            {
                string id = sheet.Id.Value;
                WorksheetPart partById = (WorksheetPart) workbookPart.GetPartById(id);
                workbookPart.DeletePart(partById);
            }
            WorksheetPart part = workbookPart.AddNewPart<WorksheetPart>();
            if (sheet == null)
            {
                uint num = (uint) (workbookPart.Workbook.Sheets.Count<OpenXmlElement>() + 1);
                Sheet sheet2 = new Sheet {
                    Id = workbookPart.GetIdOfPart(part),
                    SheetId = num,
                    Name = this.Name
                };
                workbookPart.Workbook.GetFirstChild<Sheets>().AppendChild<Sheet>(sheet = sheet2);
            }
            else
            {
                sheet.Id = workbookPart.GetIdOfPart(part);
            }
            using (OpenXmlWriter writer = OpenXmlWriter.Create(part))
            {
                this._sheetCache.WriteWorksheetPart(writer);
            }
            return true;
        }

        public ExcelCells Cells { get; protected set; }

        public ExcelColumns Columns { get; protected set; }

        public ExcelDocument Document { get; protected set; }

        internal bool Modified { get; set; }

        public string Name
        {
            get
            {
                return this._name;
            }
        }

        public ExcelRows Rows { get; protected set; }
    }
}

