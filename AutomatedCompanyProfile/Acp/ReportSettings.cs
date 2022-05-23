using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class ReportSettings
    {
        public const int MaxTotalPeers = 10;

        public readonly TimePeriodSettings TimePeriodSettings;
        public string TickerSymbol => TimePeriodSettings.TickerSymbol;
        public string ReportType; // Financials, Non-Financials, Short
        public string[] Peers;
        public string[] PeersShortNames;

        public ReportSettings()
        {
            TimePeriodSettings = new TimePeriodSettings();
        }
    }


    public class Settings
    {
        public string TickerSymbol;
        public string Error;

        public Settings()
        {
        }

        public virtual void Reset()
        {
            TickerSymbol = null;
            Error = null;
        }

        public override string ToString()
        {
            var b = new StringBuilder();
            ToString(b);
            return b.ToString();
        }

        protected virtual void ToString(StringBuilder b)
        {
            if (!string.IsNullOrEmpty(TickerSymbol))
                b.Append("TickerSymbol=").Append(TickerSymbol);

            if (!string.IsNullOrEmpty(Error))
                b.AppendIfNotEmpty("|").Append("Error=").Append(Error ?? "");
        }
    }

    public class TimePeriodSettings : Settings
    {
        public string TimePeriodType;
        public string FirstPeriod;
        public string LastPeriod;
        public string PeerFirstPeriod;
        public string PeerLastPeriod;
        public string DecompositionBegin;
        public string DecompositionEnd;
        public readonly string[] TimePeriods = new string[12];
        public readonly string[] PeerPeriods = new string[12];

        public TimePeriodSettings()
        {
        }

        public override void Reset()
        {
            base.Reset();
            TimePeriodType = null;
            DecompositionBegin = null;
            DecompositionEnd = null;
            for (var i = 0; i < TimePeriods.Length; ++i)
                TimePeriods[i] = null;
            for (var i = 0; i < PeerPeriods.Length; ++i)
                PeerPeriods[i] = null;
        }

        public override string ToString()
        {
            var b = new StringBuilder();
            ToString(b);
            return b.ToString();
        }

        protected override void ToString(StringBuilder b)
        {
            base.ToString(b);
            b.AppendIfNotEmpty("|").Append("TimePeriodType=").Append(TimePeriodType);
            if (TimePeriods[0] != null)
                b.AppendIfNotEmpty("|").Append("TimePeriods=").Append(string.Join(";", TimePeriods));
            if (PeerPeriods[0] != null)
                b.AppendIfNotEmpty("|").Append("PeerPeriods=").Append(string.Join(";", PeerPeriods));
        }
    }
}
