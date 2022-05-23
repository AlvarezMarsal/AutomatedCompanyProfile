using StockDatabase;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Acp
{
    internal partial class Client : ExcelThread
    {
        private static readonly Counter Counter = new ();
        public readonly Server Server;
        private readonly ManualResetEventSlim workingSignal;
        public bool IsWorking => workingSignal.IsSet;
        private readonly Stocks stocks;
        public readonly long Id;
        private readonly ReportSettings reportSettings;
        private double timeoutMinutes;
        private object processClientRequestLock = new object();
        private Messenger messenger;

        public Client(Server server, string workbookFilename)
            : base("Client" + Counter.Next, server.OutputFolder, workbookFilename, server)
        {
            Id = Counter.MostRecent;

            Server = server;
            workingSignal = new ManualResetEventSlim(false);

            stocks = new Stocks();
            reportSettings = new ReportSettings();

            var t = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Timeout") ?? "30";
            if (double.TryParse(t, out timeoutMinutes))
                timeoutMinutes = Math.Max(1, Math.Min(30, timeoutMinutes));
            else
                timeoutMinutes = 30;
        }

        protected override void Started()
        {
            base.Started();
        }

        protected override void Run()
        {
            try
            {
                ListenAndService();
            }
            catch (Exception ex)
            {
                Log(ex);
            }
            finally
            {
                workingSignal.Set(); // really means 'I'll clean up after myself'
            }
        }

        private void ListenAndService()
        {
            if (!ShouldKeepRunning) return;
            LoadCiqThunk();
            Log("Loaded Excel file");

            //Yield();
            if (!ShouldKeepRunning) return;
            Log("Session is ready");

            //Yield();
            messenger = new Messenger(1235, ProcessClientRequest, this);
            messenger.Start();

            while (ShouldKeepRunning)
            {
                Yield(250);
            }

            reportGenerator?.Stop();
            messenger.Stop();
            Log("Disconnected client");
        }

        protected override void AfterConnectedToAddIn()
        {
            bool ok = TryForPeriod(() =>
            {
                FastSetCellValue("CIQThunk", 47, ColumnB, "NYSE:BA");
                var value = FastGetCellValue("CIQThunk", 47, ColumnC);
                if (value != "The Boeing Company")
                    return false;

                FastSetCellValue("CIQThunk", 47, ColumnB, "NYSE:A");
                value = FastGetCellValue("CIQThunk", 47, ColumnC);
                if (value != "Agilent Technologies, Inc.")
                    return false;

                return true;

            }, 10 * 60 * 1000);

            if (!ok)
                throw new CiqInactiveException();
        }

        private string ProcessClientRequest(string request)
        {
            if (CiqErrorMessageReceived != DateTime.MinValue)
            {
                var ts = DateTime.Now - CiqErrorMessageReceived;
                if (ts.TotalMinutes > 5)
                {
                    CiqErrorMessageReceived = DateTime.MinValue;
                }
                else
                {
                    return "ERROR=The Capital IQ Service is unresponsive.";
                }
            }
            request.SplitAt('!', out var verb, out var remainder);
            var parameters = ParseParameters(remainder);

            lock (processClientRequestLock)
            {
                try
                {
                    switch (verb.ToUpper())
                    {
                        case "FINDPEERS": return ProcessFindPeers(parameters);                                      // Uses ThunkSheet, rows 25, 34-46
                        case "GENERATEREPORT": return BeginReport(parameters);
                        case "GETREPORTPROGRESS": return GetReportProgress(parameters);
                        case "QUICKTICKERSEARCH": return ProcessQuickTickerSearch(parameters);                      // Uses ThunkSheet, row 47
                        case "SEARCHBYNAME": return ProcessSearchByNameMessage(parameters);                         // Uses ThunkSheet, row 47 (via GetCompanyInfo)
                        case "VALIDATETIMEPERIODSETTINGS": return ValidateTimePeriodSettings(parameters);           // Uses ThunkSheet, rows 49, 50, 63
                        case "QUIT": return null;
                    }
                }
                catch (Exception ex)
                {
                    return "ERROR=" + ex.Message;
                }
            }

            return "ERROR=Unknown verb " + verb;
        }

        private static NamedParameters ParseParameters(string raw)
            => NamedParameters.FromString(raw, '&', '=');

        private string ProcessSearchByNameMessage(NamedParameters parameters)
        {
            try
            {
                var c = SearchByName(parameters["SearchName"]);
                return CompanyInfo.ToString(c);
            }
            catch (Exception ex)
            {
                return "ERROR=" + ex.Message;
            }
        }

        private string ProcessFindPeers(NamedParameters parameters)
        {
            var count = int.Parse(parameters["Count"]);
            var c = FindPeers(parameters["TickerSymbol"], count+1);
            return CompanyInfo.ToString(c);
        }

        private string ProcessQuickTickerSearch(NamedParameters parameters)
        {
            try
            {
                var forTarget = parameters["ForTarget"] != null;
                var c = QuickTickerSearch(parameters["SearchName"]);
                var data = c?.ToString() ?? " ";
                if ((c != null) && forTarget)
                {
                    reportSettings.TimePeriodSettings.TickerSymbol = c.TickerSymbol;
                    reportSettings.TimePeriodSettings.TimePeriodType = parameters["TimePeriodType"];
                    ValidateTimePeriodSettings();
                    data = "TimePeriodSettings=" + reportSettings.TimePeriodSettings.ToString();
                    var peers = FindPeers(c.TickerSymbol, 11);
                    data += "^Peers=" + CompanyInfo.ToString(peers);
                }
                return data;
            }
            catch (Exception ex)
            {
                return "ERROR=" + ex.Message;
            }
        }

        private string ProcessMessage<T>(NamedParameters parameters, Func<NamedParameters, T> excelHostFunction)
        {
            var results = excelHostFunction(parameters);
            return results?.ToString();
        }

        private string ProcessMessage<T>(NetworkStream stream, NamedParameters parameters, Func<NetworkStream, NamedParameters, T> excelHostFunction)
        {
            var results = excelHostFunction(stream, parameters);
            return results?.ToString();
        }

    }
}
