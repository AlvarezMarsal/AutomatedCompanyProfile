using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Acp
{
    public abstract class ExcelThread : StaThreaded
    {
        public virtual bool UsesInterop => false;
        private readonly string inputFolder;
        private readonly string rootOutputFolder;
        private readonly string outputFolder;
        protected readonly string WorkbookFilename;
        private Process excelProcess;
        protected LocalNode LocalNode;
        protected int TargetMessengerPort;
        protected Messenger Messenger { get; private set; }

        protected const int ColumnA = 1;
        protected const int ColumnB = 2;
        protected const int ColumnC = 3;
        protected const int ColumnD = 4;
        protected const int ColumnE = 5;
        protected const int ColumnF = 6;
        protected const int ColumnG = 7;
        protected const int ColumnH = 8;
        protected const int ColumnI = 9;
        protected const int ColumnJ = 10;
        protected const int ColumnK = 11;
        protected const int ColumnL = 12;
        protected const int ColumnM = 13;
        protected const int ColumnN = 14;

        public readonly Office.MsoTriState False = Office.MsoTriState.msoFalse;
        public readonly Office.MsoTriState True = Office.MsoTriState.msoTrue;

        public DateTime CiqErrorMessageReceived = DateTime.MinValue;


        public ExcelThread(string name, string outputFolder, string workbookFilename, Log log)
            : base(name, log)
        {
            Log("Creating Excel Thread " + name);
            inputFolder = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Input");
            if (!Directory.Exists(inputFolder))
                Directory.CreateDirectory(inputFolder);

            rootOutputFolder = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Output");
            if (!Directory.Exists(rootOutputFolder))
                Directory.CreateDirectory(rootOutputFolder);

            this.outputFolder = outputFolder;
            WorkbookFilename = workbookFilename;
        }

        protected override void Starting()
        {
            base.Starting();          
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("kernel32.dll", ExactSpelling = true)]
        public static extern IntPtr GetConsoleWindow();

        protected override void Started()
        {
            // Log("In ExcelThread.Started for " + Name);
            base.Started();
            SetTag("NeedsCalculate", true);

            Log("Starting EXCEL for " + Name);
            var args = "/e /r \"" + WorkbookFilename + "\"";
            for (var i = 0; i < 10; ++i)
            {
                var giveUp = DateTime.Now.AddMinutes(5);
                var ep = Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", args);
                while (DateTime.Now < giveUp)
                {
                    if (!ep.MainWindowHandle.Equals(IntPtr.Zero))
                    {
                        excelProcess = ep;
                        break;
                    }
                    Yield(1000);
                }
                if (excelProcess != null)
                    break;
            }

            if (excelProcess == null)
                throw new CiqInactiveException();

            excelProcess.WaitForInputIdle();
            SetForegroundWindow(GetConsoleWindow());

            LocalNode = new LocalNode(this);
            LocalNode.MessageReceived += (_, e) => HandleLocalNodeMessage(e);
            LocalNode.Broadcast("GETNODE", WorkbookFilename);

            if (!TryForPeriod(() => (Messenger != null), 5 * 60 * 1000))
                throw new CiqInactiveException();
        }

        override protected void RunWrapper()
        {
            base.RunWrapper();
        }

        protected override void Stopping()
        {
            Log("Stopping Excel Thread " + Name);

            LocalNode.Dispose();
            LocalNode = null;

            base.Stopping();
        }

        protected virtual string HandleLocalNodeMessage(LocalNodeMessageEventArgs e)
        {
            switch (e.Message[0].ToUpper())
            {
                case "NODEPORT":
                    if (e.Message[2] == WorkbookFilename)
                    {
                        TargetMessengerPort = int.Parse(e.Message[1]);
                        Messenger = new Messenger(HandleMessengerMessages, this);
                        Messenger.Start();
                        AfterConnectedToAddIn();
                    }
                    break;

                case "ERROR":
                    CiqErrorMessageReceived = DateTime.Now;
                    break;
            }
            return null;
        }

        protected virtual void AfterConnectedToAddIn()
        { 
        }

        protected virtual string HandleMessengerMessages(string request)
        {
            return "";
        }

        #region Set commands

        protected string[] FastSetCellValue(Excel.Worksheet sheet, int row, int column, string value)
            => FastSetCellValue(sheet.Name, row, column, value);
        protected string[] FastSetCellValue(Excel.Worksheet sheet, int row, int column, double value)
            => FastSetCellValue(sheet.Name, row, column, value);
        protected string[] FastSetCellValue(Excel.Worksheet sheet, int row, int column, int value)
            => FastSetCellValue(sheet.Name, row, column, value);

        protected string[] FastSetCellValue(string sheet, int row, int column, string value)
        {
            var b = new StringBuilder(); 
            AppendSetCellValue(b, sheet, row, column, value);
            return SendCommandsToAddIn(b);
        }

        protected string[] FastSetCellValue(string sheet, int row, int column, double value)
        {
            var b = new StringBuilder();
            AppendSetCellValue(b, sheet, row, column, value);
            return SendCommandsToAddIn(b);
        }

        protected string[] FastSetCellValue(string sheet, int row, int column, int value)
            => FastSetCellValue(sheet, row, column, (double)value);

        #endregion

        #region Get commands
        
        protected string FastGetCellValue(Excel.Worksheet sheet, int row, int column)
            => FastGetCellValue(sheet.Name, row, column);

        protected string FastGetCellValue(string sheet, int row, int column)
        {
            var b = new StringBuilder();
            AppendGetCellValue(b, sheet, row, column);
            var response = SendCommandsToAddIn(b);
            if ((response == null) || (response.Length != 1) || (response[0] == null) || response[0].StartsWith("ERROR"))
                throw new CiqInactiveException();
            return response[0];
        }

        protected string[] FastGetCellValues(string sheet, int row, int firstColumn, int lastColumn)
        {
            var b = new StringBuilder();
            AppendGetCellValue(b, sheet, row, firstColumn);
            for (var i = firstColumn + 1; i <= lastColumn; ++i)
            {
                AppendGetCellValue(b, row, i);
            }
            return SendCommandsToAddIn(b);
        }

        #endregion

        #region Build command string

        protected StringBuilder AppendSetCellValue(StringBuilder b, Excel.Worksheet sheet, int row, int column, string value)
            => AppendSetCellValue(b, sheet?.Name, row, column, value);
        protected StringBuilder AppendSetCellValue(StringBuilder b, Excel.Worksheet sheet, int row, int column, double value)
            => AppendSetCellValue(b, sheet?.Name, row, column, value);
        protected StringBuilder AppendSetCellValue(StringBuilder b, Excel.Worksheet sheet, int row, int column, int value)
            => AppendSetCellValue(b, sheet?.Name, row, column, value);

        protected StringBuilder AppendSetCellValue(StringBuilder b, int row, int column, string value)
            => AppendSetCellValue(b, (string) null, row, column, value);
        protected StringBuilder AppendSetCellValue(StringBuilder b, int row, int column, double value)
            => AppendSetCellValue(b, (string) null, row, column, value);
        protected StringBuilder AppendSetCellValue(StringBuilder b, int row, int column, int value)
            => AppendSetCellValue(b, (string) null, row, column, value);

        protected StringBuilder AppendSetCellValue(StringBuilder b, string sheet, int row, int column, string value)
        {
            if (b.Length > 0)
                b.Append('|');
            if (sheet != null)
                b.Append("USE|").Append(sheet).Append('|');
            b.Append("SET|").Append(row).Append('|').Append(column).Append('|');
            b.Append(string.IsNullOrEmpty(value) ? "@empty" : value);
            return b;
        }

        protected StringBuilder AppendSetCellValue(StringBuilder b, string sheet, int row, int column, double value)
        {
            if (b.Length > 0)
                b.Append('|');
            if (sheet != null)
                b.Append("USE|").Append(sheet).Append('|');
            b.Append("SETN|").Append(row).Append('|').Append(column).Append('|').Append(value.ToString());
            return b;
        }

        protected StringBuilder AppendSetCellValues(StringBuilder b, Excel.Worksheet sheet, int row, int column, params object[] values)
            => AppendSetCellValues(b, sheet?.Name, row, column, values);
        protected StringBuilder AppendSetCellValues(StringBuilder b, string sheet, int row, int column, params object[] values)
        {
            if (b.Length > 0)
                b.Append('|');
            if (sheet != null)
                b.Append("USE|").Append(sheet);
            foreach (var value in values)
            {
                if ((value is double) || (value is long) || (value is int))
                    b.Append("|SETN|");
                else
                    b.Append("|SET|");
                b.Append(row).Append('|').Append(column).Append('|').Append(value ?? "@empty");
                ++column;
            }
            return b;
        }


        protected StringBuilder AppendSetCellValue(StringBuilder b, string sheet, int row, int column, int value)
            => AppendSetCellValue(b, sheet, row, column, (double) value);

        protected StringBuilder AppendGetCellValue(StringBuilder b, Excel.Worksheet sheet, int row, int column)
             => AppendGetCellValue(b, sheet?.Name, row, column);
        protected StringBuilder AppendGetCellValue(StringBuilder b, int row, int column)
              => AppendGetCellValue(b, (string) null, row, column);

        protected StringBuilder AppendGetCellValue(StringBuilder b, string sheet, int row, int column)
        {
            if (b.Length > 0)
                b.Append('|');
            if (sheet != null)
                b.Append("USE|").Append(sheet).Append('|');
            b.Append("GET|").Append(row).Append('|').Append(column);
            return b;
        }

        protected StringBuilder AppendGetCellFormula(StringBuilder b, Excel.Worksheet worksheet, int row, int column)
            => AppendGetCellFormula(b, worksheet.Name, row, column);

        protected StringBuilder AppendGetCellFormula(StringBuilder b, string sheet, int row, int column)
        {
            if (b.Length > 0)
                b.Append('|');
            if (sheet != null)
                b.Append("USE|").Append(sheet).Append('|');
            return b.Append("GETFORMULA|").Append(row).Append('|').Append(column);
        }

        protected StringBuilder AppendAutomaticCalculation(StringBuilder b, bool state)
        {
            if (b.Length > 0)
                b.Append('|');
            b.Append("CALC|").Append(state ? 1 : 0);
            return b;
        }

        protected StringBuilder AppendRunScript(StringBuilder b, string name)
        {
            if (b.Length > 0)
                b.Append('|');
            b.Append("RUN|").Append(name);
            return b;
        }

        protected StringBuilder AppendActivateWorksheet(StringBuilder b, string worksheet)
        {
            if (b.Length > 0)
                b.Append('|');
            b.Append("ACTIVATE|").Append(worksheet);
            return b;
        }

        protected StringBuilder AppendUseWorksheet(StringBuilder b, string worksheet)
        {
            if (b.Length > 0)
                b.Append('|');
            b.Append("USE|").Append(worksheet);
            return b;
        }


        protected StringBuilder AppendComment(StringBuilder b, string comment, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (b.Length > 0)
                b.Append('|');
            b.Append("COMMENT|");
            if (comment != null)
                b.Append(comment).Append(" from ");
            b.Append(callerMemberName).Append(" in ").Append(callerFilePath).Append(", line ").Append(callerLineNumber);
            return b;
        }

        #endregion

        protected string[] SendCommandsToAddIn(string commands)
        {
            Log("Sending to add-in: " + commands);
            try
            {
                var response = Messenger.Send(TargetMessengerPort, commands);
                return response.Split('|');
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        protected string[] SendCommandsToAddIn(StringBuilder commands)
            => SendCommandsToAddIn(commands.ToString());

        protected bool IsErrorResponse(string[] response) => (response == null) || (response.Length == 0) || IsErrorResponse(response[0]);
        protected bool IsErrorResponse(string response) => (response == null) || response.StartsWith("ERROR");


        protected string FastGetCellFormula(Excel.Worksheet sheet, int row, int column)
            => FastGetCellFormula(sheet.Name, row, column);

        protected string FastGetCellFormula(string sheet, int row, int column)
        {
            try
            {
                var b = new StringBuilder();
                AppendGetCellFormula(b, sheet, row, column);
                var response = SendCommandsToAddIn(b);
                if (IsErrorResponse(response))
                    throw new CiqInactiveException();
                var formula = response[0];
                return formula;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        protected string FastGetCellFormula(int row, int column)
            => FastGetCellFormula((string) null, row, column);

        protected string[] AutomaticCalculation(bool state)
        {
            try
            {
                var b = new StringBuilder();
                AppendAutomaticCalculation(b, state);
                return SendCommandsToAddIn(b);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        protected string[] RunScript(string name)
        {
            try
            {
                var b = new StringBuilder();
                AppendRunScript(b, name);
                return SendCommandsToAddIn(b);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        protected string[] ActivateWorksheet(string worksheet)
        {
            try
            {
                var b = new StringBuilder();
                AppendActivateWorksheet(b, worksheet);
                return SendCommandsToAddIn(b);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        protected string[] UseWorksheet(string worksheet)
        {
            try
            {
                var b = new StringBuilder();
                AppendUseWorksheet(b, worksheet);
                return SendCommandsToAddIn(b);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

    }
}
