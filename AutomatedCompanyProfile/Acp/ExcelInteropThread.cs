using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Acp
{
    public abstract class ExcelInteropThread : ExcelThread
    {
        public override bool UsesInterop => true;
        protected Excel.Application ExcelApp;
        private Process excelProcess;
        private Excel.Workbook workbook;
        protected Excel.Sheets WorkSheets;

        public ExcelInteropThread(string name, string outputFolder, string workbookFilename, Log log) 
            : base(name, outputFolder, workbookFilename, log)
        {
        }

        protected override void Starting()
        {
            base.Starting();          
        }

        protected override void Started()
        {
            Log("In ExcelThread.Started for " + Name);
            base.Started();
            SetTag("NeedsCalculate", true);

            Log("Starting excel process for " + Name);
            var args = "/e /r \"" + WorkbookFilename + "\"";
            excelProcess = Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", args);
            var giveUp = DateTime.Now.AddMinutes(5);
            while (excelProcess.MainWindowHandle.Equals(IntPtr.Zero))
            {
                Yield(1000);
                if (DateTime.Now > giveUp)
                    throw new Exception("Excel could not start.");
            }
            while (!excelProcess.WaitForInputIdle())
            {
                if (DateTime.Now > giveUp)
                    throw new Exception("Excel could not start.");
            }
            SetForegroundWindow(GetConsoleWindow());

            Log("Creating ExcelApp for " + Name);
            ExcelApp = new Excel.Application();

            // Bring this app to the foregrouns so Excel will put itself in the
            // ROT, where we can find it.

            Log("Getting workbook for " + Name);
            ExcelApp.DisplayAlerts = false;
            //excelApp.Visible = Visible;
            //excelApp.ScreenUpdating = Visible;
            try
            {
                workbook = (Excel.Workbook)System.Runtime.InteropServices.Marshal.BindToMoniker(WorkbookFilename);
                var fn = Path.GetFileName(WorkbookFilename);
                workbook.Application.Windows[fn].Visible = true;
                workbook.Application.Visible = true;
                workbook.EnableAutoRecover = false;

                Log("Getting WorkSheets for " + Name);
                WorkSheets = workbook.Sheets;
            }
            catch (Exception ex)
            {
                Log(ex);
                Stop(0);
            }

            LocalNode = new LocalNode(this);
            LocalNode.MessageReceived += (_, e) => HandleLocalNodeMessage(e);
            bool connected = TryForPeriod(() =>
            {
                LocalNode.Broadcast("GETNODE", WorkbookFilename);
                return TryForPeriod(() => { return Messenger != null; }, 15 * 1000);

            }, 5 * 60 * 1000);

            if (!connected)
                throw new CiqInactiveException(); 
        }

        override protected void RunWrapper()
        {
            base.RunWrapper();
        }

        protected override void Stopping()
        {
            Log("Stopping ExcelThread " + Name);

            try
            {
                ExcelApp.DisplayAlerts = false;
                // ExcelApp.Run("ClearClipboard");
                ExcelApp.DisplayAlerts = false;
            }
            catch
            {
            }

            ZapComObject(ref WorkSheets);
            WorkSheets = null;

            try { workbook?.Close(false); } catch { }
            ZapComObject(ref workbook);
            workbook = null;

            try { ExcelApp?.Quit(); } catch { }
            ExcelApp = null;

            base.Stopping();
        }

        protected Excel.Worksheet FindWorksheet(string name)
        {
            for (var i = 1; i <= WorkSheets.Count; ++i)
            {
                var s = WorkSheets.Item[i];
                if (s.Name == name)
                {
                    // Log("Worksheet " + s.Name + " was found by searching for " + name);
                    return (Excel.Worksheet) s;
                }
            }

            return null;
        }

    }
}
