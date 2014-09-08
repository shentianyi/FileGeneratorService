using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace TestConsoleApplication
{
   public class OpenXMLTest
    {
       static SpreadsheetDocument document = null;
       static WorkbookPart workBook = null;
       static Worksheet worksheet = null;
        // Merge two adjacent cells in a worksheet.
        // Notice that after the merge, only the content from one cell is preserved.
      public  static void Run()
      {
            string docName = @"C:\Excel\MergeCellsEx.xlsx";
            string sheetName = "Sheet1";
            string cell1Name = "A1";
            string cell2Name = "C1";

            using( document = SpreadsheetDocument.Open(docName, true))
            {
                worksheet = GetWorksheet(sheetName);
                workBook = document.WorkbookPart;

                // Create Spreadsheet cells.
                CreateSpreadsheetCell( cell1Name);
                CreateSpreadsheetCell(cell2Name);
                MergeCells mergeCells;

                if (worksheet.Elements<MergeCells>().Count() > 0)
                    mergeCells = worksheet.Elements<MergeCells>().First();
                else
                {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Count() > 0)
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    else
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }

                // Create the merged cell and append it to the MergeCells collection.
                MergeCell mergeCell = new MergeCell()
                {
                    Reference =
                        new StringValue(cell1Name + ":" + cell2Name)
                };
                mergeCells.Append(mergeCell);
                worksheet.Save();
            }
            Console.WriteLine("The two cells are now merged.\nPress a key.");
            Console.ReadKey();
        }

      public static void RunMerge()
      {
          string docName = @"C:\Excel\my_report (7).xlsx";
          string sheetName = "First Sheet";

          using (document = SpreadsheetDocument.Open(docName, true))
          {
              worksheet = GetWorksheet(sheetName);
              workBook = document.WorkbookPart;
              MergeCells(1, 0, 3,0 );
              worksheet.Save();
          } 
      }
        // Get the specified worksheet.
        private static Worksheet GetWorksheet( string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook
                .Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart
                .GetPartById(sheets.First().Id);
            return worksheetPart.Worksheet;
        }

        // Create a spreadsheet cell. 
        private static void CreateSpreadsheetCell(string cellName)
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
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        public static Row GetRow(int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public static Cell GetCell( int rowIndex, string columnName)
        {
            Row row = GetRow( rowIndex);

            if (row == null)
                return null;
            Cell cc = row.Elements<Cell>().ToList()[0];

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }

        // get openxml cell
        internal static Cell GetOpenCell( int rowIndex, int colIndex)
        {
            Row row = GetRow( rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().ToList()[colIndex];
        }
        // get cellname
        internal static string GetCellName(int rowIndex, int colIndex)
        {
            Cell cell = GetOpenCell(rowIndex, colIndex);
            return cell == null ? null : cell.CellReference.Value;
        }

        public static void MergeCells(int row, int col, int endRow, int endCol)
        { 
                string cell1Name = GetCellName(row, col);

                string cell2Name = GetCellName(endRow, endCol);
                    MergeTwoCells(cell1Name, cell2Name); 
             
        }

        public static void MergeTwoCells(string cell1Name, string cell2Name)
        {
            MergeCells cells;
           // Worksheet worksheet = document.GetOSpreadsheet().WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First<WorksheetPart>().Worksheet;
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

    }
}
