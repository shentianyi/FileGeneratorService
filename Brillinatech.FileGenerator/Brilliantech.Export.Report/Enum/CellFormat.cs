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
                    return "##0%";
                case CellFormatType.FloatPercent:
                    return "##0.0##%";
            }
            return null;
        }



        public static object GetFormatValue(CellFormatType type, string value)
        {

            if (type.Equals(CellFormatType.None))
            {
                int iVal;
                double dVal;
                if (int.TryParse(value, out iVal))
                {
                    return iVal;
                }
                else if (double.TryParse(value, out dVal))
                {
                    return dVal;
                }
                else
                {
                    return value;
                }
            }
            else if (type.Equals(CellFormatType.IntPercent) || type.Equals(CellFormatType.FloatPercent))
            {
                return double.Parse(value.TrimEnd('%')) / 100;
            }
            return value;
        }
    }
}
