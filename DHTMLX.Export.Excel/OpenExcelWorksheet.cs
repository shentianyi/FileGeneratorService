using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace DHTMLX.Export.Excel
{
    public class OpenExcelWorksheet : IDisposable
    {
        private bool _disposed;
        public string filePath;
        public string worksheetName;

        public SpreadsheetDocument document = null;
        public WorkbookPart workBook = null;
        public Worksheet worksheet = null;
        /// <summary>
        /// write sheet
        /// </summary>
        /// <param name="path"></param>
        /// <param name="data"></param>
        public OpenExcelWorksheet(string filePath, MemoryStream data, string worksheetName)
        {
            this.filePath = filePath;
            this.worksheetName = worksheetName;
            using (data)
            {
                using (FileStream excel = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    excel.Write(data.ToArray(), 0, (int)data.Length);
                }
            }
        }

        public MemoryStream ToMS()
        {
            MemoryStream ms = new MemoryStream();
            using (FileStream excel = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                excel.CopyTo(ms);
            }
            File.Delete(filePath);
            return ms;
        }

        // Get the specified worksheet.
        private Worksheet GetWorksheet()
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook
                .Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart
                .GetPartById(sheets.First().Id);
            return worksheetPart.Worksheet;
        }
        // Create a spreadsheet cell. 
        private void CreateSpreadsheetCell(string cellName)
        {
            string columnName = GetColumnName(cellName);
            Console.WriteLine(columnName);
            uint rowIndex = 2;
            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r
                .RowIndex.Value == rowIndex);
            Row row = rows.First();
            IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference
                .Value == cellName);
        }

        // Parse the cell name to get the column name.
        private string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        public Row GetRow(int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public Cell GetCell(int rowIndex, string columnName)
        {
            Row row = GetRow(rowIndex);

            if (row == null)
                return null;
            Cell cc = row.Elements<Cell>().ToList()[0];

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }

        // get openxml cell
        internal Cell GetOpenCell(int rowIndex, int colIndex)
        {
            Row row = GetRow(rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().ToList()[colIndex];
        }
        // get cellname
        internal string GetCellName(int rowIndex, int colIndex)
        {
            Cell cell = GetOpenCell(rowIndex, colIndex);
            return cell == null ? null : cell.CellReference.Value;
        }

        public void MergeCells(int row, int col, int endRow, int endCol)
        {
            using (document = SpreadsheetDocument.Open(filePath, true))
            {
                worksheet = GetWorksheet();
                workBook = document.WorkbookPart;
                string cell1Name = GetCellName(row, col);
                string cell2Name = GetCellName(endRow, endCol);
                MergeTwoCells(cell1Name, cell2Name);
                worksheet.Save();
            }
        }

        public void MergeTwoCells(string cell1Name, string cell2Name)
        {
            MergeCells cells;
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
            MergeCell cell = new MergeCell
            {
                Reference = new StringValue(cell1Name + ":" + cell2Name)
            };
            cells.Append(new OpenXmlElement[] { cell });
        }

        ~OpenExcelWorksheet()
        {
            this.Dispose(false);
        }
        protected void Dispose(bool disposing)
        {
            if (!this._disposed)
            { 
                this._disposed = true;
            }
        }
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
