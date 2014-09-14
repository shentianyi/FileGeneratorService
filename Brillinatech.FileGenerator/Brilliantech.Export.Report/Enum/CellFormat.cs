using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Brilliantech.Export.Report.Enum
{
    public enum CellFormatType
    {
        None,
        IntPercent,
        FloatPercent
    }

    public class CellFormat
    {
        public static string GetFormatString(CellFormatType type)
        {
            switch (type)
            {
                case CellFormatType.IntPercent:
                    return "##%";
                case CellFormatType.FloatPercent:
                    return "##.##%";
            }
            return null;
        }
    }
}
