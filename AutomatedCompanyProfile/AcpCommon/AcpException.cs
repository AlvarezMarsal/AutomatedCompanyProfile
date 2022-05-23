using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class AcpException : Exception
    {
        public string CallerMemberName { get; private set; }
        public string CallerFilePath { get; private set; }
        public int CallerLineNumber { get; private set; }

        /*
        public AcpException([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base()
        {
            Setup(callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        public AcpException(string msg, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(msg)
        {
            Setup(Application.Log, callerMemberName, callerFilePath, callerLineNumber);
        }

        public AcpException(string msg, Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(msg, ex)
        {
            Setup(Application.Log, callerMemberName, callerFilePath, callerLineNumber);
        }

        public AcpException(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(ex.Message, ex)
        {
            Setup(Application.Log, callerMemberName, callerFilePath, callerLineNumber);
        }

        public AcpException(Log log, string msg, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(msg)
        {
            Setup(log, callerMemberName, callerFilePath, callerLineNumber);
        }

        public AcpException(Log log, string msg, Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(msg, ex)
        {
            Setup(log, callerMemberName, callerFilePath, callerLineNumber);
        }

        public AcpException(Log log, Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(ex.Message, ex)
        {
            Setup(log, callerMemberName, callerFilePath, callerLineNumber);
        }


        private void Setup(Log log, string callerMemberName, string callerFilePath, int callerLineNumber)
        {
            CallerMemberName = callerMemberName;
            CallerFilePath = callerFilePath;
            CallerLineNumber = callerLineNumber;
            log.Notice(GetType().Name + " thrown from " + callerMemberName + " in " + callerFilePath + ", line " + callerLineNumber);
        }


    }
}
