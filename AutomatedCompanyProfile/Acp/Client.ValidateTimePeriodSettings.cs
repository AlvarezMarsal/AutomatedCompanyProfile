using System;

namespace Acp
{
    internal partial class Client : ExcelThread
    {
        public string ValidateTimePeriodSettings(NamedParameters parameters)
        {
            SetTimePeriodSettings(parameters);
            ValidateTimePeriodSettings();
            return reportSettings.TimePeriodSettings.ToString();
        }

        private void SetTimePeriodSettings(NamedParameters parameters)
        {
            var tps = reportSettings.TimePeriodSettings;
            tps.Reset();

            foreach (var parameter in parameters)
            {
                switch (parameter.Key.ToUpper())
                {
                    case "TICKERSYMBOL":
                        tps.TickerSymbol = parameter.Value;
                        break;

                    case "TIMEPERIODTYPE":
                        tps.TimePeriodType = parameter.Value;
                        break;

                    case "FIRSTPERIOD":
                        tps.FirstPeriod = parameter.Value;
                        break;

                    case "LASTPERIOD":
                        tps.LastPeriod = parameter.Value;
                        break;

                    case "PEERFIRSTPERIOD":
                        tps.PeerFirstPeriod = parameter.Value;
                        break;

                    case "PEERLASTPERIOD":
                        tps.PeerLastPeriod = parameter.Value;
                        break;

                    case "DECOMPOSITIONEND":
                        tps.DecompositionEnd = parameter.Value;
                        break;

                    case "DECOMPOSITIONBEGIN":
                        tps.DecompositionBegin = parameter.Value;
                        break;

                    case "SENDER":
                        break;

                    case "REPORTTYPE":
                        reportSettings.ReportType = parameter.Value;
                        break;
                    case "PEERS":
                        reportSettings.Peers = parameter.Value.Split(',');
                        break;
                    case "PEERSSHORTNAMES":
                        reportSettings.PeersShortNames = parameter.Value.Split(',');
                        break;

                    default:
                        Warning("Unknown field: " + parameter.Key);
                        break;

                }
}
        }

        // Uses ThunkSheet, rows 49, 50, 63
        private void ValidateTimePeriodSettings()
        {
            Enter();

            string cps;
            string fps;

            if (string.IsNullOrEmpty(reportSettings.TickerSymbol))
            {
                Track();
                var type = reportSettings.TimePeriodSettings.TimePeriodType;
                if (type == "Quarters")
                {
                    cps = "CQ1 2020;CQ2 2020;CQ3 2020;CQ4 2020;CQ1 2021;CQ2 2021;CQ3 2021;CQ4 2021;CQ1 2022;CQ2 2022;CQ3 2022;CQ4 2022";
                    fps = "FQ2 2018;FQ3 2018;FQ4 2018;FQ1 2019;FQ2 2019;FQ3 2019;FQ4 2019;FQ1 2020;FQ2 2020;FQ3 2020;FQ4 2020;FQ1 2021";
                }
                else // if (type == "Years")
                {
                    cps = "CY2011;CY2012;CY2013;CY2014;CY2015;CY2016;CY2017;CY2018;CY2019;CY2020;CY2021;CY2022";
                    fps = "FY2010;FY2011;FY2012;FY2013;FY2014;FY2015;FY2016;FY2017;FY2018;FY2019;FY2020;FY2021";
                }
            }
            else
            {
                FastSetCellValue(thunkSheet, 49, ColumnB, reportSettings.TickerSymbol);
                if (reportSettings.TimePeriodSettings.TimePeriodType != null)
                    FastSetCellValue(thunkSheet, 50, ColumnB, reportSettings.TimePeriodSettings.TimePeriodType);

                // thunkSheet.Calculate();

                cps = FastGetCellValue(thunkSheet, 63, 3);
                fps = FastGetCellValue(thunkSheet, 63, 6);
            }

            if (fps != null)
            {
                var periods = fps.Split(';');
                for (var i = 0; i < periods.Length; ++i)
                    reportSettings.TimePeriodSettings.TimePeriods[i] = periods[i];
            }

            if (cps != null)
            {
                var periods = cps.Split(';');
                for (var i = 0; i < periods.Length; ++i)
                    reportSettings.TimePeriodSettings.PeerPeriods[i] = periods[i];
            }

            Exit();
        }
    }
}
