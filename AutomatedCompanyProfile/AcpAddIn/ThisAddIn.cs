using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace AcpAddIn
{
    public partial class ThisAddIn
    {
        Acp.LogUser log;
        Acp.Messenger messenger;
        Acp.LocalNode localNode;
        IntPtr[] pumpMessageHandles;
        readonly System.Reflection.Missing Missing = System.Reflection.Missing.Value;

        Excel.Workbook _workbook = null;
        Excel.Workbook Workbook
        {
            get
            {
                if (_workbook == null)
                {
                    var ok = TryForPeriod(() =>
                    {
                        _workbook = Application.ActiveWorkbook;
                        return (_workbook != null);

                    });
                }

                return _workbook;
            }
        }

        Excel.Sheets _worksheets = null;
        Excel.Sheets Worksheets
        {
            get
            {
                if (_worksheets == null)
                {
                    var giveUp = DateTime.Now.AddMinutes(1);
                    while ((_worksheets == null) && (DateTime.Now < giveUp))
                    {
                        try { _worksheets = Workbook.Worksheets; } catch { }
                        PumpMessagesUntilReady();
                    }
                    if (_worksheets == null)
                        throw new Exception();
                }

                return _worksheets;
            }
        }

        Excel.Worksheet CurrentWorksheet = null;

        Acp.Signal CiqErrorSignal;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var neverSetEvent = new ManualResetEvent(false);
            pumpMessageHandles = new IntPtr[] { neverSetEvent.SafeWaitHandle.DangerousGetHandle() };

            log = new Acp.LogUser(new Acp.FileLog(null, "AcpAddIn"));
            messenger = new Acp.Messenger(HandleRequest, log);
            messenger.Start();

            localNode = new Acp.LocalNode(log);
            localNode.MessageReceived += (_, e) => HandleLocalNodeMessage(e);

            CiqErrorSignal = Acp.Application.Signals.Get("CiqError", true);

            Application.EnableAnimations = false;
            Application.EnableLargeOperationAlert = false;
            Application.EnableSound = false;
            var a = Application.AutoRecover;
            a.Time = 120;
            Application.DisplayAlerts = false;

            log.Log("Started");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            log.Log("Stopping");
            localNode.Dispose();
            messenger.Stop();
            log.Log("Stopped");
        }

        [DllImport("ole32.dll")]
        static extern int CoWaitForMultipleHandles(uint dwFlags, uint dwTimeout, uint cHandles, IntPtr[] pHandles, out uint lpdwindex);

        private void PumpMessages()
        {
            CoWaitForMultipleHandles(0, 0, (uint)pumpMessageHandles.Length, pumpMessageHandles, out var _);
        }
        private void PumpMessagesUntilReady()
        {
            bool ready = false;
            while (!ready)
            { 
                PumpMessages(); 
                try { ready = Application.Ready; } catch { }
            }
        }

        private string HandleRequest(string request)
        {
            try
            {
                var parts = request.Split('|');
                var i = 0;
                var b = new StringBuilder();

                PumpMessagesUntilReady();

                do
                {
                    if (parts[i] == "SET")
                    {
                        int row = int.Parse(parts[++i]);
                        int column = int.Parse(parts[++i]);
                        var value = parts[++i];
                        if (value == "@empty")
                            value = "";

                        if (!Set(row, column, value))
                            throw new Acp.CiqInactiveException();
                    }
                    else if (parts[i] == "SETN")
                    {
                        int row = int.Parse(parts[++i]);
                        int column = int.Parse(parts[++i]);
                        var value = double.Parse(parts[++i]);

                        if (!SetNumber(row, column, value))
                            throw new Acp.CiqInactiveException();
                    }
                    else if (parts[i] == "GET")
                    {
                        int row = int.Parse(parts[++i]);
                        int column = int.Parse(parts[++i]);
                        var vstr = Get(row, column);
                        if (vstr == null)
                            throw new Acp.CiqInactiveException();
                        if (b.Length > 0)
                            b.Append('|');
                        b.Append(vstr);
                    }
                    else if (parts[i] == "GETFORMULA")
                    {
                        int row = int.Parse(parts[++i]);
                        int column = int.Parse(parts[++i]);

                        PumpMessagesUntilReady();

                        var cells = GetWorksheetCells(CurrentWorksheet);
                        var obj = cells.Formula;
                        log.Log("Formula " + (obj ?? "is null"));
                        if (b.Length > 0)
                            b.Append('|');
                        b.Append(obj?.ToString() ?? "");
                    }
                    else if (parts[i] == "CALC")
                    {
                        int state = int.Parse(parts[++i]);
                        if (state == 0)
                        {
                            Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                        }
                        else
                        {
                            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                            Application.CalculateFull();
                            log.Log("Full recalc complete");
                        }
                    }
                    else if (parts[i] == "RUN")
                    {
                        var script = parts[++i];

                        Application.Run(script, Missing, Missing, Missing, Missing,
                            Missing, Missing, Missing, Missing, Missing, Missing,
                            Missing, Missing, Missing, Missing, Missing, Missing,
                            Missing, Missing, Missing, Missing, Missing, Missing,
                            Missing, Missing, Missing, Missing, Missing, Missing,
                            Missing, Missing);

                    }
                    else if (parts[i] == "ACTIVATE")
                    {
                        var name = parts[++i];
                        if (!UseWorksheet(name))
                            break;
                        Workbook.Activate();
                        CurrentWorksheet.Activate();

                        var success = TryForPeriod(() => Application.ActiveSheet.Name == name);

                        if (!success)
                        {
                            log.Log("Total ACTIVATE failure");
                            throw new Acp.CiqInactiveException();
                        }
                    }
                    else if (parts[i] == "USE")
                    {
                        var name = parts[++i];
                        log.Log("USE " + name);

                        if (!UseWorksheet(name))
                            throw new Exception("Worksheet '" + "' could not be used");
                        
                        Workbook.Activate();
                        CurrentWorksheet.Activate();

                        var success = TryForPeriod(() => Application.ActiveSheet.Name == name);
                        if (!success)
                            throw new Acp.CiqInactiveException();
                    }
                    else if (parts[i] == "COMMENT")
                    {
                        log.Info(parts[++i]);
                    }
                    else
                    {
                        return "ERROR=Unknown command " + parts[i];
                    }

                } while (++i < parts.Length);

                return (b.Length == 0) ? " " : b.ToString();
            }
            catch (Acp.CiqInactiveException)
            {
                return "ERROR=" + Acp.Application.CiqSucks;
            }
            catch (Exception ex)
            {
                return "ERROR=" + ex.ToString();
            }
        }

        private bool Set(int row, int column, object value)
        {
            log.Log("SET " + row + " " + column + " " + value.ToString());

            var success = TryForPeriod(() =>
            {
                var cells = GetWorksheetCells(CurrentWorksheet);
                var cell = (Excel.Range) cells[row, column];
                if (cell == null)
                    return false;

                cell.Formula = "";
                cell.Value2 = value;

                var verify = GetOnce(row, column);
                bool ok = value.Equals(verify);
                if (!ok)
                    log.Log("SET failure: " + verify.ToString() + " vs " + value.ToString());
                //else
                //    log.Trace("Values matched, SET worked");
                return ok;

            }, 15 * 60 * 1000, 5000);

            if (!success)
                log.Log("Total SET failure");

            return success;
        }

        private bool SetNumber(int row, int column, double value)
        {
            log.Log("SETN " + row + " " + column + " " + value.ToString());

            var success = TryForPeriod(() =>
            {
                var cells = GetWorksheetCells(CurrentWorksheet);
                var cell = (Excel.Range)cells[row, column];
                if (cell == null)
                    return false;

                cell.Formula = "";
                cell.Value2 = value;

                var verify = GetOnce(row, column);
                if (verify == null)
                    return false;

                if (verify is double d)
                {
                    var e = Math.Abs(d - value);
                    return e < 0.000001;
                }
                else
                {
                    log.Log("SETN failure: " + verify.GetType().Name + ": " + verify.ToString() + " vs " + value.ToString());
                    return false;
                }
            });

            if (!success)
                log.Log("Total SETN failure");

            return success;
        }


        private object Get(int row, int column)
        {
            log.Trace("GET " + row + " " + column);

            object value = null;
            var ok = TryForPeriod(() =>
            {
                value = GetOnce(row, column);
                return (value != null);

            }, 15 * 60 * 1000, 5000);

            if (value == null)
                log.Log("GET " + row + " " + column + " returning null");
            return value;
        }

        private object GetOnce(int row, int column)
        {
            log.Log("GETONCE " + row + " " + column);
            try
            {
                var cells = GetWorksheetCells(CurrentWorksheet);
                var cell = cells[row, column];
                if (cell == null)
                    return false;

                var value = (object) cell.Value2;

                if (value is null)
                {
                    log.Log("Value is null");
                    return "";
                }

                if (value is string s)
                {
                    if (s.StartsWith("#")) // error
                    {
                        log.Log("Value is ERROR " + s);
                        return null;
                    }
                    log.Log("Value is string '" + s + "'");
                    return s;
                }

                if (value is long l)
                {
                    log.Log("Value is long " + l);
                    return (double) l;
                }

                if (value is double d)
                {
                    log.Log("Value is double " + d);
                    return d;
                }

                log.Log("Bad type: " + value.GetType());
                s = value.ToString();
                log.Log("Value is " + value.GetType() + ": " + s);
                if (s.StartsWith("#")) // error
                    return null;
                return s;
            }
            catch (Exception ex)
            {
                log.Log(ex);
                return null;
            }
        }

        private bool UseWorksheet(string name)
        {
            if (name == CurrentWorksheet?.Name)
                return true;

            log.Log("Switching to worksheet '" + name + "' from '" + (CurrentWorksheet?.Name ?? "null") + "'");
            CurrentWorksheet = null;

            foreach (Excel.Worksheet w in Worksheets)
            {
                if (w.Name == name)
                {
                    CurrentWorksheet = w;
                    log.Log("Worksheet '" + name + "' found");
                    return true;
                }

                PumpMessagesUntilReady();
            }

            log.Warning("Worksheet '" + name + "' not found.");
            return false;
        }

        private void HandleLocalNodeMessage(Acp.LocalNodeMessageEventArgs e)
        {
            PumpMessagesUntilReady();

            log.Log(e.ToString());
            switch (e.Message[0].ToUpper())
            {
                case "GETNODE":
                    {
                        var filename = e.Message[1];
                        if (filename != null)
                        {
                            string wbfn = GetWorkbookFullName(Workbook);
                            if (string.Compare(filename, wbfn, true) == 0)
                            {
                                log.Log("Responding");
                                e.Respond("NODEPORT", messenger.Port.ToString(), filename);
                            }
                        }
                    }
                    break;

                default:
                    log.Warning(e.ToString());
                    break;
            }
        }

        private string GetWorkbookFullName(Excel.Workbook workbook)
            => TryForPeriod(() => workbook.FullName);

        private Excel.Range GetWorksheetCells(Excel.Worksheet worksheet)
            => TryForPeriod(() => worksheet.Cells);

        private bool TryForPeriod(Func<bool> func, int totalTime = 15 * 60 * 1000, int yieldTime = 1000)
        {
            var now = DateTime.Now;
            var giveUp = now.AddMilliseconds(totalTime);
            do
            {
                try
                {
                    var done = func();
                    if (done)
                        return true;
                }
                catch (Exception ex)
                {
                    log.Trace(ex);
                }

                try
                {
                    var pump = now.AddMilliseconds(yieldTime);
                    do
                    {
                        PumpMessagesUntilReady();
                        now = DateTime.Now;

                    } while (now < pump);
                }
                catch (Exception ex)
                {
                    log.Trace(ex);
                }

            } while (now < giveUp);

            log.Log("Giving up");
            return false;
        }

        private T TryForPeriod<T>(Func<T> func, int totalTime = 15 * 60 * 1000, int yieldTime = 1000)
            where T : class
        {
            T value = null;
            var ok = TryForPeriod(() =>
            {
                value = func();
                return (value != null);

            }, totalTime, yieldTime);
            return value;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
