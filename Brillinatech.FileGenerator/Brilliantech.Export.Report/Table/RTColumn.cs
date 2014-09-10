using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Xml;

namespace Brilliantech.Export.Report.Table
{
    public class RTColumn
    {
        private string colName;
        private string type;
        private string align;
        private int colspan;
        private int rowspan;
        private int width = 100;
        private int height = 1;
        private bool is_footer = false;
        private Color backgroundColor = ColorTranslator.FromHtml("#D1E5FE");
        private ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin;
        private Color borderColor = ColorTranslator.FromHtml("#A4BED4");
      

        public RTColumn()
        {
             
        }

        public RTColumn(XmlElement parent)
        {
            is_footer = parent.ParentNode.ParentNode.Name.Equals("foot");

            colName = parent.HasChildNodes ? parent.FirstChild.Value : "";
            if (parent.HasAttribute("width"))
            {
                width = int.Parse(parent.Attributes["width"].Value);
            }
            type = parent.GetAttribute("type");
            align = parent.GetAttribute("align");

            if (parent.HasAttribute("colspan"))
            {
                colspan = int.Parse(parent.Attributes["colspan"].Value);
            }

            if (parent.HasAttribute("rowspan"))
            {
                rowspan = int.Parse(parent.Attributes["rowspan"].Value);
            }
        }

        public string ColName
        {
            get { return colName; }
            set { colName = value; }
        }
        public string Type
        {
            get { return type; }
            set { type = value; }
        }
        public string Align
        {
            get { return align; }
            set { align = value; }
        }
        public int Colspan
        {
            get { return colspan; }
            set { colspan = value; }
        }
        public int Rowspan
        {
            get { return rowspan; }
            set { rowspan = value; }
        }
        public int Width
        {
            get { return width; }
            set { width = value; }
        }
        public int Height
        {
            get { return height; }
            set { height = value; }
        }
        public bool Is_footer
        {
            get { return is_footer; }
            set { is_footer = value; }
        }
        //public int RowIndex { get; set; }
        //public bool IsHead { get; set; }
        //public string Value { get; set; }
        public Color BackgroundColor
        {
            get { return backgroundColor; }
            set { backgroundColor = value; }
        }

        public ExcelBorderStyle BorderStyle
        {
            get { return borderStyle; }
            set { borderStyle = value; }
        }
        public Color BorderColor
        {
            get { return borderColor; }
            set { borderColor = value; }
        }
    }
}
