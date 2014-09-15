using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using Brilliantech.Export.Report.Enum;

namespace Brilliantech.Export.Report.Table
{
    public class RTCell
    {
        private string value = "";

        private ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin;
        private Color borderColor = ColorTranslator.FromHtml("#A4BED4");
        private string cellFormatString = null;
        private CellFormatType cellFormatType = CellFormatType.None;
        private Color backgroundColor;
        private string backgroundColorString;

        public RTCell() { }

        public RTCell(XmlNode parent)
        {
            XmlElement ele = (XmlElement)parent;
            if (ele.HasAttribute("value"))
            {
                this.value = parent.Attributes["value"].Value;
            }
            if (parent.HasChildNodes && this.value == null)
            {
                value = parent.FirstChild.Value;
            }
           
            if (ele.HasAttribute("format")) {
                this.cellFormatType = (CellFormatType)int.Parse(parent.Attributes["format"].Value);
                this.cellFormatString = CellFormat.GetFormatString(this.cellFormatType);
            }
            if (ele.HasAttribute("bgcolor"))
            {
                try
                {
                    this.backgroundColorString = parent.Attributes["bgcolor"].Value.TrimStart('#');
                    this.backgroundColor = ColorTranslator.FromHtml("#" + this.backgroundColorString);
                }
                catch {
                    this.backgroundColorString = null;
                }
            }
        }


        public string Value
        {
            get { return this.value; }
        }

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

        public Color GetBackgroundColor(int row)
        {
            if (this.backgroundColorString != null)
                return this.backgroundColor;
            return row % 2 == 0 ? ColorTranslator.FromHtml("#FFFFFF") : ColorTranslator.FromHtml("#E3EFFF");
        }

        public object GetValue()
        {
            //int iVal;
            //double dVal;
            //if (int.TryParse(this.value, out iVal))
            //{
            //    return iVal;
            //}
            //else if (double.TryParse(this.value, out dVal))
            //{
            //    return dVal;
            //}
            //return this.value;
            return CellFormat.GetFormatValue(this.cellFormatType, this.value);
        }


        public string CellFormatString
        {
            get { return cellFormatString; }
            set { cellFormatString = value; }
        }

        public CellFormatType CellFormatType
        {
            get { return cellFormatType; }
            set { cellFormatType = value; }
        }
    }
}
