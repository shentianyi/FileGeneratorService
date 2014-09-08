namespace OpenExcel.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;

    public static class BinarySearchExtension
    {
        public static int BinarySearch<T>(this IList<T> list, T value) where T: IComparable<T>
        {
            return list.BinarySearch<T>(value, Comparer<T>.Default);
        }

        public static int BinarySearch<T>(this IList<T> list, T value, IComparer<T> comparer)
        {
            return list.BinarySearch<T>(0, list.Count, value, comparer);
        }

        public static int BinarySearch<T>(this IList<T> list, int index, int length, T value, IComparer<T> comparer)
        {
            if (list == null)
            {
                throw new ArgumentNullException("list");
            }
            if ((index < 0) || (length < 0))
            {
                throw new ArgumentOutOfRangeException((index < 0) ? "index" : "length");
            }
            if ((list.Count - index) < length)
            {
                throw new ArgumentException();
            }
            int num = index;
            int num2 = (index + length) - 1;
            while (num <= num2)
            {
                int num3 = num + ((num2 - num) >> 1);
                int num4 = comparer.Compare(list[num3], value);
                if (num4 == 0)
                {
                    return num3;
                }
                if (num4 < 0)
                {
                    num = num3 + 1;
                }
                else
                {
                    num2 = num3 - 1;
                }
            }
            return ~num;
        }
    }
}

