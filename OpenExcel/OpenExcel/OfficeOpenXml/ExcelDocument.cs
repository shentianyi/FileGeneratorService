namespace OpenExcel.OfficeOpenXml
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using OpenExcel.Common;
    using OpenExcel.OfficeOpenXml.Internal;
    using OpenExcel.OfficeOpenXml.Style;
    using System;
    using System.IO;
    using System.IO.Packaging;
    using System.Runtime.CompilerServices;

    public class ExcelDocument : IDisposable
    {
        private bool _disposed;
        private SpreadsheetDocument _doc;
        private DocumentSharedStrings _sharedStrings;
        private DocumentStyles _styles;

        private ExcelDocument(SpreadsheetDocument doc)
        {
            this._doc = doc;
            WorkbookPart workbookPart = this.GetOSpreadsheet().WorkbookPart;
            this._styles = new DocumentStyles(workbookPart);
            this._sharedStrings = new DocumentSharedStrings(workbookPart);
            this.Workbook = new ExcelWorkbook(this);
        }

        private void Cleanup()
        {
            bool flag = false;
            foreach (ExcelWorksheet worksheet in this.Workbook.Worksheets)
            {
                if (worksheet.Save())
                {
                    flag = true;
                }
            }
            WorkbookPart workbookPart = this.GetOSpreadsheet().WorkbookPart;
            if (flag && (workbookPart.CalculationChainPart != null))
            {
                workbookPart.DeletePart(workbookPart.CalculationChainPart);
            }
        }

        private static ExcelDocument CreateBlankWorkbook(SpreadsheetDocument doc)
        {
            doc.AddWorkbookPart();
            doc.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            return new ExcelDocument(doc);
        }

        public ExcelFont CreateFont(string name, double size)
        {
            return new ExcelFont(null, this._styles, null) { Name = name, Size = size };
        }

        public static ExcelDocument CreateWorkbook(Package package)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(package, workbook));
        }

        public static ExcelDocument CreateWorkbook(Stream stream)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(stream, workbook));
        }

        public static ExcelDocument CreateWorkbook(string path)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(path, workbook));
        }

        public static ExcelDocument CreateWorkbook(Package package, bool autoSave)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(package, workbook, autoSave));
        }

        public static ExcelDocument CreateWorkbook(Stream stream, bool autoSave)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(stream, workbook, autoSave));
        }

        public static ExcelDocument CreateWorkbook(string path, bool autoSave)
        {
            SpreadsheetDocumentType workbook = SpreadsheetDocumentType.Workbook;
            return CreateBlankWorkbook(SpreadsheetDocument.Create(path, workbook, autoSave));
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    this.Cleanup();
                    this.SharedStrings.Save();
                }
                this._doc.Dispose();
                this._disposed = true;
            }
        }

        public void EnsureStylesDefined()
        {
            this._styles.EnsureStylesheet();
        }

        ~ExcelDocument()
        {
            this.Dispose(false);
        }

        internal SpreadsheetDocument GetOSpreadsheet()
        {
            return this._doc;
        }

        public static ExcelDocument Open(Package package)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(package));
        }

        public static ExcelDocument Open(Package package, OpenSettings openSettings)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(package, openSettings));
        }

        public static ExcelDocument Open(Stream stream, bool isEditable)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(stream, isEditable));
        }

        public static ExcelDocument Open(string path, bool isEditable)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(path, isEditable));
        }

        public static ExcelDocument Open(Stream stream, bool isEditable, OpenSettings openSettings)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(stream, isEditable, openSettings));
        }

        public static ExcelDocument Open(string path, bool isEditable, OpenSettings openSettings)
        {
            return new ExcelDocument(SpreadsheetDocument.Open(path, isEditable, openSettings));
        }

        internal void RecalcCellReferences(SheetChange sheetChange)
        {
            foreach (ExcelWorksheet worksheet in this.Workbook.Worksheets)
            {
                worksheet.RecalcCellReferences(sheetChange);
            }
        }

        internal DocumentSharedStrings SharedStrings
        {
            get
            {
                return this._sharedStrings;
            }
        }

        public DocumentStyles Styles
        {
            get
            {
                return this._styles;
            }
        }

        public ExcelWorkbook Workbook { get; protected set; }
    }
}

