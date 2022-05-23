using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public static class Application
    {
        private static string exeFileName;
        public static string ExeFileName = exeFileName ??= GetExeFileName();

        private static string exeName;
        public static string ExeName = exeName ??= Path.GetFileNameWithoutExtension(ExeFileName);

        private static string publisher;
        public static string Publisher = publisher ??= "AlvarezMarsal";

        private static string name;
        public static string Name
        {
            get => name ?? ExeName;
            set
            {
                name = string.IsNullOrWhiteSpace(value) ? null : value;
            }
        }

        private static Log log = new NullLog();
        public static Log Log
        {
            get => log;
            set
            {
                if (log != value)
                {
                    log.Dispose();
                    log = (value ?? new NullLog());
                }
            }
        }

        public static string CiqSucks = "The Capital IQ Service is not responding.";

        private static string GetExeFileName()
        {
            var ass = Assembly.GetEntryAssembly();
            if (ass != null)
                return ass.Location;
            return System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
        }

        private static void SetRegistryValue(string name, object value, RegistryValueKind kind)
        {
            var subkeyName = @"SOFTWARE\" + Application.Publisher;
            var subkey = Registry.LocalMachine.OpenSubKey(subkeyName, true) ?? Registry.LocalMachine.CreateSubKey(subkeyName);
            subkeyName += @"\" + Application.Name;
            subkey = Registry.LocalMachine.OpenSubKey(subkeyName, true) ?? Registry.LocalMachine.CreateSubKey(subkeyName);
            subkey.SetValue(name, value, kind);
        }

        public static void SetRegistryValue(string name, int value)
            => SetRegistryValue(name, value, RegistryValueKind.DWord);
        public static void SetRegistryValue(string name, long value)
            => SetRegistryValue(name, value, RegistryValueKind.QWord);
        public static void SetRegistryValue(string name, string value)
            => SetRegistryValue(name, value ?? "", RegistryValueKind.String);
        public static void SetRegistryValue(string name, bool value)
            => SetRegistryValue(name, value ? 1 : 0, RegistryValueKind.DWord);

        private static object GetRegistryValue(string name, out RegistryValueKind kind)
        {
            var subkeyName = @"SOFTWARE\" + Application.Publisher + @"\" + Application.Name;
            var subkey = Registry.LocalMachine.OpenSubKey(subkeyName, true);
            if (subkey == null)
            {
                kind = RegistryValueKind.None;
                return null;
            }
            kind = subkey.GetValueKind(name);
            return subkey.GetValue(name, null);
        }

        public static bool GetRegistryValue(string name, ref bool value)
        {
            long l = 0;
            var ok = GetRegistryValue(name, ref l);
            if (!ok)
                return false;
            value = (l != 0);
            return true;
        }

        public static bool GetRegistryValue(string name, ref int value)
        {
            long l = 0;
            var ok = GetRegistryValue(name, ref l);
            if (!ok)
                return false;
            if ((l < int.MinValue) || (l > int.MaxValue))
                return false;
            value = (int) l;
            return true;
        }

        public static bool GetRegistryValue(string name, ref long value)
        {
            var v = GetRegistryValue(name, out var kind);
            if (v == null)
                return false;

            switch (kind)
            {
                case RegistryValueKind.DWord:
                    value = (long)((int) v);
                    return true;

                case RegistryValueKind.String:
                    if (long.TryParse((string)v, out var l))
                    {
                        value = l;
                        return true;
                    }
                    return false;

                case RegistryValueKind.QWord:
                    value = (long)v;
                    return true;
            }

            return false;
        }

        public static bool GetRegistryValue(string name, ref string value)
        {
            var v = GetRegistryValue(name, out var kind);
            if (v == null)
                return false;

            switch (kind)
            {
                case RegistryValueKind.None:
                case RegistryValueKind.Unknown:
                    return false;

                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                    value = (string)v;
                    return true;

                case RegistryValueKind.Binary:
                case RegistryValueKind.MultiString:
                    return false;

                case RegistryValueKind.DWord:
                case RegistryValueKind.QWord:
                    value = v.ToString();
                    return true;
            }

            return false;
        }

        public static SignalCollection Signals = new SignalCollection();
    }
}
