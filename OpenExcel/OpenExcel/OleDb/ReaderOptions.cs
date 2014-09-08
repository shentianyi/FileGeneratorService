namespace OpenExcel.OleDb
{
    using System;
    using System.Runtime.CompilerServices;

    public class ReaderOptions
    {
        public ReaderOptions()
        {
            this.IMEX = IMEXOptions.No;
        }

        public IMEXOptions IMEX { get; set; }
    }
}

