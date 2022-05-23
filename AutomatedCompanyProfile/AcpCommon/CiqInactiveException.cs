using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class CiqInactiveException : AcpException
    {
        public CiqInactiveException([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            : base(Application.CiqSucks, callerMemberName, callerFilePath, callerLineNumber)
        {
            SetRegistryFlag(true);
            Application.Signals.Set("CiqError");
        }

        public static void SetRegistryFlag(bool state)
            => Application.SetRegistryValue("CiqInactiveException", state);
 
        public static bool GetRegistryFlag()
        {
            bool value = false;
            var ok = Application.GetRegistryValue("CiqInactiveException", ref value);
            return ok && value;
        }
    }
}
