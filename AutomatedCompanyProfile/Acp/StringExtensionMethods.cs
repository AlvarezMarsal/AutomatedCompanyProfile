using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Acp
{
    public static class StringExtensionMethods
    {
        public static int Compare(this string a, string b) => string.Compare(a, b, true);
        public static int CompareI(this string a, string b) => string.Compare(a, b, true);
        public static int CompareS(this string a, string b) => string.Compare(a, b, false);
        public static string Quoted(this string s) => "\"" + (s ?? "") + "\"";

        public static string[] SplitAt(this string s, char e)
        {
            var index = s.IndexOf(e);
            if (index < 0)
                return new string[] { s, "" };
            return new string[] { s.Substring(0, index), s.Substring(index + 1) };
        }

        public static bool SplitAt(this string s, char e, out string before, out string after)
        {
            var index = s.IndexOf(e);

            if (index <= 0)
            {
                before = s;
                after = "";
                return (index == 0);

            }

            if (index == (s.Length - 1))
            {
                before = "";
                after = s;
                return true; 
            }

            before = s.Substring(0, index);
            after = s.Substring(index+1);
            return true;
        }

    }
}
