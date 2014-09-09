// namespace OpenExcel.OfficeOpenXml.Style
//{
//    using DocumentFormat.OpenXml;
//    using DocumentFormat.OpenXml.Spreadsheet;
//    using OpenExcel.OfficeOpenXml.Internal;
//    using System;
//    using System.Runtime.CompilerServices;

//    public class ExcelAlignment
//    {
//       // private uint? _aligmentId;
//        private Alignment _alignment;
//        private IStylable _stylable;
//        private DocumentStyles _styles;

//        internal ExcelAlignment(IStylable stylable, DocumentStyles styles, Alignment? alignment) {
//            this._stylable = stylable;
//            this._styles = styles;
//           // this._aligmentId = aligmentId;
//            //if (this._aligmentId.HasValue)
//            //{
//            //    this.AlignmentObject = (Alignment)this._styles.GetAligment(this._aligmentId.Value).CloneNode(true);
//            //}
//            //else {
//            //    this.AlignmentObject = new Alignment();
//            //}
//            if (alignment.HasValue)
//            {
//                this.AlignmentObject = alignment.Value;
//            }
//            else {
//                this.AlignmentObject = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true };
//            }

//        }

//        internal Alignment AlignmentObject { get; set; }
//    }
//}
