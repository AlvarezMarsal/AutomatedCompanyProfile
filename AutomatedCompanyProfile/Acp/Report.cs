using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class Report
    {
        public Report() { }
        public string Filename { get; internal set; }
        public List<CompanyInfo> Companies { get; private set; } = new List<CompanyInfo>();
        public CompanyInfo Target => Companies[0];
    }
}
