using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class CompanyInfo
    {
        public string TickerSymbol;
        public string Name;
        public string ShortName;
        public string MarketCap;
        public string Revenue;
        public string NetProfit;
        public string Employees;

        public override string ToString()
        {
            var b = new StringBuilder();
            AppendString(b);
            return b.ToString();
        }

        public static string ToString(CompanyInfo[] infos)
        {
            var b = new StringBuilder();
            for (int i = 0; i < infos.Length; ++i)
            {
                if ((infos[i] != null) && !string.IsNullOrWhiteSpace(infos[i].TickerSymbol))
                {
                    b.AppendIfNotEmpty("|");
                    infos[i].AppendString(b);
                }
            }

            return b.ToString();
        }

        private void AppendString(StringBuilder b)
        {
            b.Append(TickerSymbol).Append(";");
            b.Append(Name).Append(";");
            b.Append(ShortName).Append(";");

            if (double.TryParse(MarketCap, out var d))
                b.Append(string.Format("{0:0,0}", d)).Append(";");
            else
                b.Append(MarketCap).Append(";");

            if (double.TryParse(Revenue, out d))
                b.Append(string.Format("{0:0,0}", d)).Append(";");
            else
                b.Append(Revenue).Append(";");

            if (double.TryParse(NetProfit, out d))
                b.Append(string.Format("{0:0,0}", d)).Append(";");
            else
                b.Append(NetProfit).Append(";");

            b.Append(Employees);
        }
    }
}
