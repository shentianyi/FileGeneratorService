using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Brilliantech.Export.Report
{
    public class RTCell
    {
        public RTCell()
        {
            this.IsHead = false;
            this.BorderStyle = ExcelBorderStyle.Thin;
            this.BorderColor = Color.Black;
        }
        
        public int RowIndex { get; set; }
        public bool IsHead { get; set; }
        public string Value { get; set; }
        public ExcelBorderStyle BorderStyle { get; set; }
        public Color BorderColor { get; set; }
    }
}
