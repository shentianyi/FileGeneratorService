using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Brilliantech.Export.Report.Enum
{
    public enum ChartAxisFormatType
    {
        None,
        IntPercent,
        FloatPercent
    }

    public class ChartAxisFormat
    {
        public static string GetFormatString(ChartAxisFormatType type)
        {
            switch (type)
            {
                case ChartAxisFormatType.IntPercent:
                    return "##%";
                case ChartAxisFormatType.FloatPercent:
                    return "##.##%";
            }
            return null;
        }
    }
}
