using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Brilliantech.Export.Report.Table
{
    public class RTRow
    {
        private RTCell[] cells;

        public RTRow() { }

        public RTRow(XmlNode parent)
        {
            XmlNodeList nodes = ((XmlElement)parent).GetElementsByTagName("cell");

            if (nodes != null && nodes.Count > 0)
            {
                cells = new RTCell[nodes.Count];
                for (int i = 0; i < nodes.Count; i++)
                {
                    cells[i] = nodes[i] == null ? new RTCell() : new RTCell(nodes[i]);
                }
            }
        }

        public RTCell[] Cells
        {
            get { return cells; }
        }
    }
}
