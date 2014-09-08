namespace OpenExcel.Utilities
{
    using System;

    public static class ValueChecker
    {
        public static bool IsNumeric(Type valueType)
        {
            TypeCode typeCode = Type.GetTypeCode(valueType);
            if ((((typeCode != TypeCode.Int16) && (typeCode != TypeCode.Int32)) && ((typeCode != TypeCode.Int64) && (typeCode != TypeCode.UInt16))) && (((typeCode != TypeCode.UInt32) && (typeCode != TypeCode.UInt64)) && ((typeCode != TypeCode.Double) && (typeCode != TypeCode.Decimal))))
            {
                return false;
            }
            return true;
        }
    }
}

