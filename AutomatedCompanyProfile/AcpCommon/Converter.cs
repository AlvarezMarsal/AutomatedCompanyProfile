using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public static class Converter
    {
        public static string ToAcpString(object o)
        {
            if (o == null)
                return "@null";
            if (o is bool b)
                return b ? "@true" : "@false";
            if (o is string s)
                return s;
            if (o is double d)
                return "!" + d.ToString();
            if (o is float f)
                return "!" + f.ToString();
            if (o is int  i)
                return "!" + i.ToString();
            if (o is long n)
                return "!" + n.ToString();
            return o.ToString();
        }

        // There are 5 data types: 
        // Number
        // Text
        // Logical
        // Error
        public static object FromAcpString(string s)
        {
            if (string.IsNullOrWhiteSpace(s) || (s == "@null"))
                return "";
            if (s.StartsWith("#"))
                return s; // ?
            if (s == "@true")
                return true;
            if (s == "@false")
                return false;
            if ((s.Length > 1) && (s[0] == '!') && double.TryParse(s.Substring(1), out var number))
                return number;
            return s;
        }

        /*
        public static bool FromAcpString(string o, out object p)
        {
            p = FromAcpString(o);
            return true;
        }

        public static bool FromString<T>(string o, out T p)
        {
            var q = FromAcpString(o);
            if (q is T b)
            {
                p = b;
                return true;
            }
            p = default;
            return false;
        }
        */
    }
}
