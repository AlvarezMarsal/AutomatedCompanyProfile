using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockDatabase
{
    public partial class Stocks
    {
        public Stocks()
        {
            Setup1();
            Setup2();
            Setup3();
            Setup4();
        }

        public string QuickFind(string input, bool fallbackToSlow = true)
        {
            input = input.Trim().ToUpper();

            if (input.Contains(':'))
            {
                int index = Array.BinarySearch<string>(LongSymbols, input);
                if (index < 0)
                    return null;
                return input;
            }

            int isi = InitialShortSymbolCharacters.IndexOf(input[0]);
            if (isi >= 0)
            {
                if (ShortSymbols[isi].TryGetValue(input, out var exchangeIndexes))
                    return ExchangeCodes[exchangeIndexes[0]] + ":" + input;
            }

            if (fallbackToSlow)
            {
                input = input.Replace(' ', ';');
                var matches = FindMatches(input, null, null, 1);
                if (matches.Length == 1)
                    return matches[0];
            }

            return null;
        }

        public IEnumerable<string> FindMatches(string input, string[] exchanges, string[] regions)
        {
            if (input == null)
                yield break;

            input = input.Trim().ToUpper().Replace(' ', ';');
            if (input.Length == 0)
                yield break;

            var colon = input.IndexOf(':');
            if (colon >= 0)
            {
                if (exchanges == null)
                {
                    exchanges = new string[1];
                    exchanges[0] = input.Substring(0, colon);
                }
                input = input.Substring(colon + 1);
            }

            var test = ";" + input + ";";
            var pos = 0;
            while (pos < VeryLongString.Length)
            {
                var match = VeryLongString.IndexOf(test, pos);
                if (match < 0)
                    break;
                var m = CheckExchangesAndRegions(match, exchanges, regions);
                if (m != null)
                    yield return m;
                pos = match + 1;
            }

            test = ";" + input;
            pos = 0;
            while (pos < VeryLongString.Length)
            {
                var match = VeryLongString.IndexOf(test, pos);
                if (match < 0)
                    break;
                var m = CheckExchangesAndRegions(match, exchanges, regions);
                if (m != null)
                    yield return m;
                pos = match + 1;
            }

            test = input;
            pos = VeryLongString.IndexOf(';');
            while (pos < VeryLongString.Length)
            {
                var end = VeryLongString.IndexOf('~', pos);
                var match = VeryLongString.IndexOf(test, pos, end - pos);
                if (match < 0)
                    break;
                var m = CheckExchangesAndRegions(match, exchanges, regions);
                if (m != null)
                    yield return m;
                pos = end - 1;
            }
        }
 
        public string[] FindMatches(string input, string[] exchanges, string[] regions, int max)
        {
            var list = new List<string>(max);
            foreach (var match in FindMatches(input, exchanges, regions))
            {
                list.Add(match);
                if (list.Count == max)
                    break;
            }
            return list.ToArray();
        }

        private string CheckExchangesAndRegions(int position, string[] exchanges, string[] regions)
        {
            var lineStart = VeryLongString.LastIndexOf('~', position);
            var lineEnd = VeryLongString.IndexOf('~', position);
            if ((lineStart < 0) || (lineEnd < 0))
                return null;

            if (exchanges != null)
            {
            }

            if (regions != null)
            {
            }

            var symbolEnd = VeryLongString.IndexOf('!', lineStart);
            var symbol = VeryLongString.Substring(lineStart + 1, symbolEnd - lineStart - 1);
            return symbol;
        }
    }
}
