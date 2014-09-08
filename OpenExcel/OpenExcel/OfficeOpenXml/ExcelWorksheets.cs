namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Threading;
    using DocumentFormat.OpenXml;

    public class ExcelWorksheets : IEnumerable<ExcelWorksheet>, IEnumerable
    {
        private Dictionary<string, ExcelWorksheet> _sheets = new Dictionary<string, ExcelWorksheet>();
        protected Worksheet worksheet;

        internal ExcelWorksheets(ExcelDocument parent)
        {
            this.Document = parent;
        }

        public ExcelWorksheet Add(string sheetName)
        {
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            if ((from s in workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>()
                where s.Name == sheetName
                select s).Count<Sheet>() > 0)
            {
                throw new InvalidOperationException("Sheet \"" + sheetName + "\" already exists.");
            }
            uint num = (uint) (workbookPart.Workbook.Sheets.Count<OpenXmlElement>() + 1);
            WorksheetPart part = workbookPart.AddNewPart<WorksheetPart>();
            Sheet newChild = new Sheet {
                Id = workbookPart.GetIdOfPart(part),
                SheetId = num,
                Name = sheetName
            };
            workbookPart.Workbook.GetFirstChild<Sheets>().AppendChild<Sheet>(newChild);
            part.Worksheet = new Worksheet();
            this.worksheet = part.Worksheet;
            part.Worksheet.Append(new OpenXmlElement[] { new SheetData() });
            part.Worksheet.Save();
            return this[sheetName];
        }

        private IEnumerable<ExcelWorksheet> EnumerateWorksheets()
        {
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            foreach (Sheet iteratorVariable1 in workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>())
            {
                yield return this[(string) iteratorVariable1.Name];
            }
        }

        public IEnumerator<ExcelWorksheet> GetEnumerator()
        {
            foreach (ExcelWorksheet iteratorVariable0 in this.EnumerateWorksheets())
            {
                yield return iteratorVariable0;
            }
        }

        public void MoveAfter(string sheetName, string referenceSheetName)
        {
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            Sheets sheets = workbookPart.Workbook.Sheets;
            Sheet newChild = (from s in sheets.Elements<Sheet>()
                where s.Name == sheetName
                select s).First<Sheet>();
            if (newChild == null)
            {
                throw new InvalidOperationException("Sheet \"" + sheetName + "\" does not exist.");
            }
            Sheet refChild = (from s in workbookPart.Workbook.Sheets.Elements<Sheet>()
                where s.Name == referenceSheetName
                select s).First<Sheet>();
            if (refChild == null)
            {
                throw new InvalidOperationException("Sheet \"" + referenceSheetName + "\" does not exist.");
            }
            newChild.Remove();
            sheets.InsertAfter<Sheet>(newChild, refChild);
        }

        public void MoveToEnd(string sheetName)
        {
            Sheets sheets = this.Document.GetOSpreadsheet().WorkbookPart.Workbook.Sheets;
            Sheet sheet = (from s in sheets.Elements<Sheet>()
                where s.Name == sheetName
                select s).First<Sheet>();
            if (sheet == null)
            {
                throw new InvalidOperationException("Sheet \"" + sheetName + "\" does not exist.");
            }
            sheet.Remove();
            sheets.Append(new OpenXmlElement[] { sheet });
        }

        public void Remove(string sheetName)
        {
            WorkbookPart workbookPart = this.Document.GetOSpreadsheet().WorkbookPart;
            Sheet sheet = (from s in workbookPart.Workbook.Sheets.Elements<Sheet>()
                where s.Name == sheetName
                select s).First<Sheet>();
            if (sheet == null)
            {
                throw new InvalidOperationException("Sheet \"" + sheetName + "\" does not exist.");
            }
            string id = sheet.Id.Value;
            WorksheetPart partById = (WorksheetPart) workbookPart.GetPartById(id);
            workbookPart.DeletePart(partById);
            sheet.Remove();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (ExcelWorksheet iteratorVariable0 in this.EnumerateWorksheets())
            {
                yield return iteratorVariable0;
            }
        }

        public ExcelDocument Document { get; protected set; }

        public ExcelWorksheet this[string name]
        {
            get
            {
                ExcelWorksheet worksheet;
                if (!this._sheets.TryGetValue(name, out worksheet))
                {
                    this._sheets[name] = worksheet = new ExcelWorksheet(name, this.Document);
                }
                return worksheet;
            }
        }



    }
}

