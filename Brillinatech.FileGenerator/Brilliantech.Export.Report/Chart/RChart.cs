using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Xml;

namespace Brilliantech.Export.Report.Chart
{
    public class RChart
    {
        private string title;
        private int height;
        private int width;

        public RChart() { }

        public RChart(XmlElement parent)
        {
            title = parent.GetAttribute("title");
            height = int.Parse(parent.GetAttribute("height"));
            width = int.Parse(parent.GetAttribute("width"));
        }

        public string Title
        {
            get { return title; }
            set { title = value; }
        }
        public int Height
        {
            get { return height; }
            set { height = value; }
        }
        public int Width
        {
            get { return width; }
            set { width = value; }
        }
    }
}
