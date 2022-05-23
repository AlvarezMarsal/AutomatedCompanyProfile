using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading;


namespace Acp
{
    internal partial class Client : ExcelThread
    {
        ReportGenerator reportGenerator;

        internal string BeginReport(NamedParameters parameters)
        {
            string response = null;

            SetTimePeriodSettings(parameters);

            foreach (var parameter in parameters)
            {
            }

            Info("TickerSymbol = " + reportSettings.TickerSymbol);
            if (!IsValidTickerSymbol(reportSettings.TickerSymbol, out var _))
            {
                response = "ERROR=Invalid ticker symbol";
                Log("Ticker symbol " + reportSettings.TickerSymbol + " is not valid");
            }
            else
            {
                var excelWorkingFilename = Server.MapOutputFilename(Application.ExeName + "-" + Id + ".xlsm");
                var templateFilename = Server.MapInputFilename(ReportGenerator.ExcelTemplateFilename);
                File.Copy(templateFilename, excelWorkingFilename, true);
                var a = File.GetAttributes(excelWorkingFilename);
                File.SetAttributes(excelWorkingFilename, a & ~System.IO.FileAttributes.ReadOnly);
                Info("Excel file is " + excelWorkingFilename);

                reportGenerator = new ReportGenerator(this, reportSettings, excelWorkingFilename);
                reportGenerator.Start();
            }

            response ??= GetReportProgress(null);
            return response;
        }


        internal string GetReportProgress(NamedParameters parameters)
        {
            var rg = reportGenerator;
            if (rg == null)
                return "";

            rg.GetProgress(out var currentStep, out var maxStep, out var filename, out var done);
            var response = "COUNT=" + currentStep + ";TOTAL=" + maxStep;
            if (!string.IsNullOrEmpty(filename))
                response += ";FILENAME=" + filename.Replace("\\", "/");
            if (done)
            {
                response += ";DONE=1";
                reportGenerator.Dispose();
                reportGenerator = null;
            }
            return response;
        }
    }
}
