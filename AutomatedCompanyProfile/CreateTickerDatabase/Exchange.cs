using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace CreateTickerDatabase
{
    class Exchange
    {
        public readonly string Code;
        public string Name;
        public int Group = int.MaxValue;     // by default, sort to the bottom
        public int StockCount;
        public int Index;   // the index at runtime

        public override string ToString()
        {
            return Code;
        }

        public Exchange(string code)
        {
            //if (code == "NASDAQ")
            //    Debug.WriteLine("oops");
            Code = code;
        }
    }
}