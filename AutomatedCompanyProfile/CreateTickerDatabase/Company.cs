using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace CreateTickerDatabase
{
    class Company
    {
        public string Name;
        private Dictionary<string, string> tickersByExchangeCode = new Dictionary<string, string>();
        public SortedSet<string> Regions = new SortedSet<string>();
        public int LongSymbolIndex;
        public SortedList<int, string> TickersByExchangeIndex;
        private string firstExchangeCode;
        private string firstTickerCode;

        public Company(string exchangeCode, string ticker, string name)
        {
            Name = name;
            tickersByExchangeCode.Add(firstExchangeCode = exchangeCode, firstTickerCode = ticker);
        }

        public void AddExchange(string exchangeCode, string ticker)
        {
            if (!tickersByExchangeCode.ContainsKey(exchangeCode))
                tickersByExchangeCode.Add(exchangeCode, ticker);
            //           else if (tickersByExchangeCode[exchangeCode] != ticker)
            //               throw new Exception();
        }

        public void SortExchanges(List<Exchange> exchanges)
        {
            TickersByExchangeIndex = new SortedList<int, string>();
            foreach (var tbx in tickersByExchangeCode)
            {
                bool found = true;
                for (var i = 0; i < exchanges.Count; ++i)
                {
                    if (exchanges[i].Code == tbx.Key)
                    {
                        TickersByExchangeIndex.Add(i, tbx.Value);
                        found = true;
                        break;
                    }
                }
                if (!found)
                    throw new Exception();
            }
        }

        public override string ToString()
        {
            return Name + " (" + firstExchangeCode + ":" + firstTickerCode + ")";
        }
    }
}