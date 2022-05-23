using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
//using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Threading;
using System.IO;
using System.Text;

namespace Acp
{
    internal partial class Client : ExcelThread
    {
        string thunkSheet = "CIQThunk";
        static readonly object companyInfoCacheLock = new object();
        static readonly SortedList<string, CompanyInfo> companyInfoCache = new SortedList<string, CompanyInfo>();

        #region Loading

        private void LoadCiqThunk()
        {
            Enter();

            Exit();
        }

        #endregion

        public enum CIQPeriod
        {
            IQ_FY   = 1000,
            IQ_FQ   = 500,
        }

        protected override void Stopping()
        {
            if (UsesInterop)
            {
                ZapComObject(ref thunkSheet);
                thunkSheet = null;
            }

            base.Stopping();

            try { File.Delete(WorkbookFilename); } catch { }
        }

        private bool IsValidTickerSymbol(string tickerSymbol, out string errorMessage)
        {
            Enter();
            bool success = false;
            string value = null;
            for (var i = 0; i < 10; ++i)
            {
                Track();
                value = CIQ(tickerSymbol, "IQ_PERIODDATE_BS", CIQPeriod.IQ_FY);
                if (value == null)
                    throw new CiqInactiveException();
                success = double.TryParse(value, out var j);
                Info("CIQ returned '" + value + "' = " + j);
                if (success)
                    break;
                Yield(1000);
            }
            errorMessage = "";
            Exit();
            return true;
        }

        // Uses thunkSheet, row 47
        private CompanyInfo GetCompanyInfo(string tickerSymbol)
        {
            Info("GetCompanyInfo(\"" + tickerSymbol + "\")");
            CompanyInfo info;

            lock (companyInfoCacheLock)
            {
                if (companyInfoCache.TryGetValue(tickerSymbol, out info))
                {
                    Debug.Assert(info.TickerSymbol == tickerSymbol);
                    Exit();
                    return info;
                }
            }

            FastSetCellValue(thunkSheet, 47, 2, tickerSymbol);
            //thunkSheet.Calculate();

            Track();
            var commands = new StringBuilder();
            AppendGetCellValue(commands, thunkSheet, 47, ColumnC);
            for (var i=ColumnD; i<=ColumnG; ++i)
                AppendGetCellValue(commands, 47, i);
            var cells = SendCommandsToAddIn(commands);
            info = new CompanyInfo { TickerSymbol = tickerSymbol };
            info.Name = cells[0];
            info.ShortName = GetCompanyShortName(info.Name);
            info.MarketCap = cells[1];
            info.Revenue = cells[2];
            info.NetProfit = cells[3];
            info.Employees = cells[4];

            lock (companyInfoCacheLock)
            {
                companyInfoCache[info.TickerSymbol] = info;
            }

            return info;
        }

        // Uses thunkSheet, row 47 (via GetCompanyInfo)
        internal CompanyInfo[] SearchByName(string name, int count = 5)
        {
            Enter(name);
            name = name.Trim().ToUpper().Replace(" ", "|");
            var alreadyFound = new HashSet<string>();
            var stats = new CompanyInfo[count];
            var index = 0;
            foreach (var found in stocks.FindMatches(name, null, null))
            {
                if (!alreadyFound.Contains(found))
                {
                    stats[index] = GetCompanyInfo(found);
                    if (stats[index] != null)
                    {
                        if (alreadyFound.Add(stats[index].TickerSymbol))
                        {
                            if (++index == stats.Length)
                                break;
                        }
                    }
                }
            }

            while (index < stats.Length)
            {
                stats[index] = stats[index] ?? new CompanyInfo();
                stats[index++].TickerSymbol = "";
            }

            Exit();
            return stats;
        }

        internal CompanyInfo QuickTickerSearch(string searchFor)
        {
            Enter();
            var ticker = stocks.QuickFind(searchFor);
            if (ticker != null)
            {
                Exit();
                return GetCompanyInfo(ticker);
            }
            Exit();
            return null;
        }

        // Uses thunkSheet, rows 25, 34-46
        internal CompanyInfo[] FindPeers(string proposedTickerSymbol, int count)
        {
            var stats = new CompanyInfo[count];

            FastSetCellValue(thunkSheet, 25, 2, proposedTickerSymbol);
            // thunkSheet.Calculate();
            //Yield();

            while (true)
            {
                var cells = FastGetCellValues(thunkSheet, 26, ColumnC, ColumnH);
                stats[0] = stats[0] ?? new CompanyInfo();
                stats[0].TickerSymbol = cells[0];
                stats[0].Name = cells[1];
                stats[0].MarketCap = cells[2];
                stats[0].Revenue = cells[3];
                stats[0].NetProfit = cells[4];
                stats[0].Employees = cells[5];
                stats[0].ShortName = GetCompanyShortName(stats[0].Name);
                var retry = (stats[0].Name == "#REFRESH") || (stats[0].Name == "(Invalid Identifier)");

                for (int i = 1, row = 35; i < stats.Length; ++i, ++row)
                {
                    cells = FastGetCellValues(thunkSheet, row, ColumnC, ColumnH);
                    stats[i] = stats[i] ?? new CompanyInfo();
                    stats[i].TickerSymbol = cells[0];
                    stats[i].Name = cells[1];
                    stats[i].MarketCap = cells[2];
                    stats[i].Revenue = cells[3];
                    stats[i].NetProfit = cells[4];
                    stats[i].Employees = cells[5];
                    stats[i].ShortName = GetCompanyShortName(stats[i].Name);
                    if ((stats[i].Name == "#REFRESH") || (stats[i].Name == "(Invalid Identifier)"))
                        retry = true;
                }

                if (!retry)
                    break;

                // Yield();
 
            }

            return stats;
        }

        internal  string GetCompanyShortName(string name)
        {
            if (name.StartsWith("The ", StringComparison.CurrentCultureIgnoreCase))
                name = name.Substring(4);
            if (name.StartsWith("Organización ", StringComparison.CurrentCultureIgnoreCase))
                name = name.Substring(13);
            var parts = name.Split(new char[] { ' ', ',', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
                return parts[0];
            if (parts[0].Length >= 4)
                return parts[0];
            if (parts[0].Length < 10)
                return parts[0] + " " + parts[1];
            return parts[0];
        }

        private string CIQ(string tickerSymbol, string dataPoint, CIQPeriod period)
        {
            var commands = new StringBuilder();
            AppendComment(commands, "In CIQ");
            AppendSetCellValues(commands, thunkSheet, 1, ColumnB, tickerSymbol, dataPoint, (int)period, "");
            AppendGetCellValue(commands, 1, ColumnA);

            var response = SendCommandsToAddIn(commands);
            if (IsErrorResponse(response))
                return null;

            return response[0];
        }

        /*
        private string CIQ(string tickerSymbol, string dataPoint, CIQPeriod period, DateTime asOf)
        {
            var d = $"{asOf.Month}/{asOf.Day}/{asOf.Year}";

            thunkSheet.SetCellValue(1, 2, tickerSymbol, this);
            thunkSheet.SetCellValue(2, 2, dataPoint, this);
            thunkSheet.SetCellValue(3, 2, (int)period, this);
            thunkSheet.SetCellValue(4, 2, d, this);
            thunkSheet.Calculate();

            var c = thunkSheet.GetCellText(1, 3, this);

            // _ciqThunkBook.Save();

            return c ?? "";
        }
        */
        /*
        private void CIQSearchByName(string searchFor, CompanyInfo[] stats)
        {
            thunkSheet.SetCellValue(11, 2, searchFor, this);
            thunkSheet.Calculate();

            var j = 0;
            for (int i = 0, row = 12; i < stats.Length; ++i, ++row)
            {
                var ticker = thunkSheet.GetCellText(row, 4, this);
                if (!string.IsNullOrWhiteSpace(ticker) && (ticker != "NA") && (ticker != "(Invalid Identifier)"))
                {
                    stats[j] = stats[j] ?? new CompanyInfo();
                    stats[j].TickerSymbol = ticker;
                    stats[j].Name = thunkSheet.GetCellText(row, 5, this);
                    stats[j].MarketCap = thunkSheet.GetCellFloatText(row, 6, this);
                    stats[j].Revenue = thunkSheet.GetCellFloatText(row, 7, this);
                    stats[j].NetProfit = thunkSheet.GetCellFloatText(row, 8, this);
                    stats[j].Employees = thunkSheet.GetCellFloatText(row, 9, this);
                    stats[j].ShortName = GetCompanyShortName(stats[j].Name);
                    ++j;
                }
            }

            while (j < stats.Length)
            {
                stats[j] = stats[j] ?? new CompanyInfo();
                stats[j].TickerSymbol = "";
                stats[j].Name = "";
                stats[j].MarketCap = "";
                stats[j].Revenue = "";
                stats[j].NetProfit = "";
                stats[j].Employees = "";
                stats[j].ShortName = "";
                ++j;
            }
        }
        */
    }
}
