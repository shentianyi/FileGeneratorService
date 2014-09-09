using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Brilliantech.Export.Report
{
    public class RTRow
    {
        public List<RTCell> Cells { get; set; }

        public void AddCells(List<RTCell> cells) {
            this.Cells = cells;
        }
    }
}
