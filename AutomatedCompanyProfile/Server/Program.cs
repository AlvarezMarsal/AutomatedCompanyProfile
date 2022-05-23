using Acp;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Server
{
    static class Program
    {
        public static int Pid;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            Thread.CurrentThread.Name = "Main";
            Pid = Process.GetCurrentProcess().Id;
            Console.Title = "Server";

            bool restart = true;
            var ciqErrorSignal = Acp.Application.Signals.Get("CiqError", true);
            while (restart)
            {
                ciqErrorSignal.Reset();

                var rootOutputFolder = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Output");
                var setting = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-ReinstallCiq");
                var reinstallCiq = false;
                if (setting == "yes")
                    reinstallCiq = true;
                else if (setting == "conditional")
                    reinstallCiq = CiqInactiveException.GetRegistryFlag();

                using (var log = new FileLog(rootOutputFolder, "Server"))
                {
                    Acp.Application.Log = log;
                    Console.Title = Path.GetFileNameWithoutExtension(log.LogFileName);

                    KillExcel(log);
                    KillPowerpoint(log);
                    FixAcpAddIn(log);
                    if (reinstallCiq)
                        ReinstallCiq(log);
                    EnableCiq(log);
                    InstallAcpAddIn(log);

                    var app = new App(log);
                    restart = app.Run(ciqErrorSignal);
                    Acp.Application.Log = null;
                }
            }

            Environment.Exit(0);
        }

        private static void FixAcpAddIn(Log log)
        {
            RegistryKey acpAddIn;

            try
            {
                acpAddIn = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\AcpAddIn", true);
            }
            catch
            {
                acpAddIn = null;
            }

            if (acpAddIn == null)
            {
                try
                {
                    acpAddIn = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\AcpAddIn");
                }
                catch
                {
                    return;
                }
            }

            acpAddIn.SetValue("Description", "AcpAddIn");
            acpAddIn.SetValue("FriendlyName", "AcpAddIn");
            acpAddIn.SetValue("LoadBehavior", 3);
            acpAddIn.SetValue("Manifest", @"C:\Program Files (x86)\AlvarezAndMarsal\AcpAddIn\AcpAddIn.vsto|vstolocal");

            try
            {
                acpAddIn = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel", true);
            }
            catch
            {
                acpAddIn = null;
            }

            if (acpAddIn == null)
            {
                try
                {
                    acpAddIn = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel");
                }
                catch
                {
                    return;
                }
            }

            acpAddIn.SetValue("MyComputer", "Enabled");
        }

        private static bool KillExcel(Log log)
            => KillProcessByName("EXCEL", log);
        private static bool KillPowerpoint(Log log)
             => KillProcessByName("POWERPNT", log);

        private static void EnableCiq(Log log)
        {
            var ciqAddIn = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\Addins\CIQAddin.Connect", true);
            if (ciqAddIn == null)
                return;
            ciqAddIn.SetValue("LoadBehavior", 3);

            var disabled = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\excel\Resiliency\DisabledItems", true);
            if (disabled != null)
            {
                var names = disabled.GetValueNames();
                foreach (var name in names)
                    try { disabled.DeleteValue(name); } catch (Exception ex) { Debug.WriteLine(ex); };
            }
        }

        private static bool KillProcessByName(string name, Log log)
        {
            log.Write("Killing " + name + " processes");
            bool killedAny = false;
 
            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName.Equals(name))
                {
                    try
                    {
                        process.Kill();
                        killedAny = true;
                    }
                    catch (Exception ex)
                    {
                        log.Write("Could not kill " + name);
                        log.Write(ex);
                    }
                }

                // Debug.WriteLine(process.ProcessName);
                // Debug.WriteLine("    " + (process.MainWindowTitle ?? null));
            }

            return killedAny;
        }


        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr hWndChildAfter, string className, string windowTitle);

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        static readonly string CiqInstallerName = "CapitalIQ_OfficeBootstrap.exe";

        private static void ReinstallCiq(Log log)
        {
            KillProcessByName(Path.GetFileNameWithoutExtension(CiqInstallerName), log);

            var ass = Assembly.GetEntryAssembly();
            var exe = ass.Location;
            var installer = Path.Combine(Path.GetDirectoryName(exe), CiqInstallerName);
            if (File.Exists(installer))
            {
                var psi = new ProcessStartInfo();
                psi.UseShellExecute = true;
                psi.FileName = installer;
                psi.Arguments = "/u /s";
                var process = Process.Start(psi);

                var giveup = DateTime.Now.AddMinutes(1);
                while (process.MainWindowHandle == IntPtr.Zero)
                {
                    if (DateTime.Now > giveup)
                    {
                        process.Kill();
                        return;
                    }
                    Thread.Sleep(1000);
                }

                int keystrokeDelay = 500;
                SetForegroundWindow(process.MainWindowHandle);
                SendKeys.SendWait("{ENTER}");

                Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait(" {TAB}{TAB}{ENTER}"); // agree to terms
                /* Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{TAB}");
                Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{TAB}");
                Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{ENTER}"); */

                Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB} "); // no power point

                //Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{TAB}{TAB}{TAB}{TAB} "); // no word
                //Thread.Sleep(keystrokeDelay);

                Thread.Sleep(keystrokeDelay);
                SendKeys.SendWait("{TAB}{ENTER}");  // install!

                var timeOut = DateTime.Now.AddSeconds(60);
                while (DateTime.Now < timeOut)
                {
                    if (process.HasExited)
                        break;
                    Thread.Sleep(1000);
                }

                if (!process.HasExited)
                {
                    SendKeys.SendWait("{ENTER}");   // finish
                    process.WaitForExit();
                }

                process.WaitForExit();
                CiqInactiveException.SetRegistryFlag(false);
            }
        }
        private static void InstallAcpAddIn(Log log)
        {
            var ass = Assembly.GetEntryAssembly();
            var exe = ass.Location;
            var installer = Path.Combine(Path.GetDirectoryName(exe), "AcpAddInSetup.msi");
            if (File.Exists(installer))
            {
                var psi = new ProcessStartInfo();
                psi.UseShellExecute = true;
                psi.FileName = installer;
                psi.Arguments = "/passive";
                var process = Process.Start(psi);

                process.WaitForExit();
                var timeOut = DateTime.Now.AddSeconds(60);
                while (DateTime.Now < timeOut)
                {
                    if (process.HasExited)
                        break;
                    Thread.Sleep(1000);
                }
            }
        }
    }

    class App : LogUser
    {
        public App(Log log) : base(log)
        {
        }

        public bool Run(Acp.Signal ciqErrorSignal)
        {
            // Start listening.  If we can't, there's already a copy of the program running
            // (or some other program using the same port).
            bool restart = true;
            using (var server = new Acp.Server(this))
            {
                server.TryToCleanCache = true;

                var logLevel = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-LogLevel");
                if (!string.IsNullOrWhiteSpace(logLevel))
                {
                    var c = char.ToUpper(logLevel.Trim()[0]);
                    server.Settings.SetLevel(c);
                }

                try
                {
                    server.Start();

                    Console.WriteLine("-----------------------------------------");
                    Console.WriteLine("Hit the ESC key to terminate this program");
                    Console.WriteLine("-----------------------------------------");

                    while (!server.IsStopped)
                    {
                        if (ciqErrorSignal.IsSet)
                            break;

                        if (Console.KeyAvailable)
                        {
                            var k = Console.ReadKey(true);
                            if (k.Key == ConsoleKey.Escape)
                            {
                                Console.WriteLine("User has requested a shutdown");
                                restart = false;
                                break;
                            }
                        }
                    }

                    server.Stop(Timeout.Infinite);
                }
                catch (Exception ex)
                {
                    Log(ex);
                    restart = false;
                }
            }
            return restart;
        }
    }
}
