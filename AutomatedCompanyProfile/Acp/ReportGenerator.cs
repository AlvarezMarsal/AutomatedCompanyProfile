using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Runtime.CompilerServices;

namespace Acp
{
    class ReportGenerator : ExcelInteropThread
    {
        bool disposed;
        const bool UseOldPowerPointGenerationCode = false;
        const bool CopyGraphsAsPictures = false;
        readonly ReportSettings reportSettings;
        PowerPoint.Application powerpointApp;
        string inputWorkSheet;
        string impDataWorkSheet;
        string exportNWorkSheet;
        string exportFWorkSheet;
        string exportWorkSheet;
        public const string ExcelTemplateFilename = "Template.xlsm";
        private const string PowerPointTemplateFilename = "Template.pptx";
        private readonly Client client;
        Report report;
        private readonly object progressLock = new object();
        private int currentStep;
        private int maxStep = 307;
        private bool done = false;
        private string currentMemberName;
        private int currentLineNumber;
        private ManualResetEvent connectedSignal;
        private ManualResetEvent failedSignal;
        private WaitHandle[] signals;

        //public string ReportFilename => report?.Filename;
        //public int CurrentStep { get { lock (progressLock) { return currentStep; } } }
        //public int MaxStep { get { lock (progressLock) { return maxStep; } } }
        //public bool Done { get { lock (progressLock) { return maxStep; } } }

        public ReportGenerator(Client client, ReportSettings reportSettings, string workbookFilename) 
            : base("ReportGenerator", client.Server.OutputFolder, workbookFilename, client)
        {
            this.client = client;
            this.reportSettings = reportSettings;

            connectedSignal = new ManualResetEvent(false);
            failedSignal = new ManualResetEvent(false);
            signals = new WaitHandle[] { connectedSignal, failedSignal };
        }

        #region Progress Tracking

    public void GetProgress(out int currentStep, out int maxStep, out string filename, out bool done)
        {
            lock (progressLock)
            {
                currentStep = this.currentStep;
                maxStep = this.maxStep;
                filename = report?.Filename;
                done = this.done;
                Debug.WriteLine(currentMemberName + ":" + currentLineNumber);
            }
        }

        private int IncrementCurrentStep([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
            => IncrementCurrentStep(1, callerMemberName, callerFilePath, callerLineNumber);

        private int IncrementCurrentStep(int n, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!ShouldKeepRunning)
                throw new AcpException("Exception");

            lock (progressLock)
            {
                Debug.WriteLine((currentMemberName = callerMemberName) + ":" + (currentLineNumber = callerLineNumber));

                currentStep += n;
                if (currentStep >= maxStep)
                    maxStep = currentStep + 1;
                return currentStep;
            }
        }


        private int SetCurrentStep(int n, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            lock (progressLock)
            {
                Debug.WriteLine((currentMemberName = callerMemberName) + ":" + (currentLineNumber = callerLineNumber));
                currentStep = n;
                if (currentStep >= maxStep)
                    maxStep = currentStep + 1;
                return currentStep;
            }
        }

        #endregion Progress Tracking

        #region Loading

        private void LoadPowerPoint()
        {
            powerpointApp = new PowerPoint.Application();
            if (powerpointApp == null)
                throw new Exception("Could not start PowerPoint");
            IncrementCurrentStep();
        }

        private void LoadExcelTemplate()
        {
            IncrementCurrentStep();

            // Open the working Excel file
/*
            templateWorkbook = workbooks.Open(
                excelWorkingFilename,                      // filename
                Type.Missing,                               // UpdateLinks
                false,                                      // Readonly
                Type.Missing,                               // Format
                "1234",                                     // Password
                "1234",                                     // WriteResPassword
                true,                                       // IgnoreReadOnlyRecommended
                Type.Missing,                               // Origin
                Type.Missing,                               // Delimiter
                true,                                       // Editable
                false,                                      // Notify
                Type.Missing,                               // Converter
                false,                                      // AddToMru
                Type.Missing,                               // Local
                Excel.XlCorruptLoad.xlNormalLoad);          // CorruptLoad
*/

            inputWorkSheet = "Input";
            impDataWorkSheet = "ImpData";
            exportNWorkSheet = "Export (N)";
            exportFWorkSheet = "Export (F)";
        }

        #endregion

        protected override void Run()
        {
            lock (progressLock)
            {
                currentStep = 0;
                done = false;
            }

            LoadPowerPoint();
            LoadExcelTemplate();

            report = new Report();
            IncrementCurrentStep();

            var s = WaitHandle.WaitAny(signals, 5 * 60 * 1000);
            if (s != 0)
                throw new CiqInactiveException();

            Log("Before SetSpreadsheetInputs");

            AutomaticCalculation(false);
            SetSpreadsheetInputs();
            AutomaticCalculation(true);

            Log("After SetSpreadsheetInputs");

            DoPeerGeneration();

            /*
                        DoExcelGeneration();

                        if (UseOldPowerPointGenerationCode)
                        {
            #pragma warning disable CS0162 // Unreachable code detected
                            DoPowerPointGeneration();
            #pragma warning restore CS0162 // Unreachable code detected
                        }
                        else
                        {
            #pragma warning disable CS0162 // Unreachable code detected
                            DoBetterPowerPointGeneration();
            #pragma warning restore CS0162 // Unreachable code detected
                        }
            */

            lock (progressLock)
                done = true;

            UnloadOfficeApps();
        }

        protected override void AfterConnectedToAddIn()
        {
            bool ok = TryForPeriod(() =>
            {
                FastSetCellValue("Input", 5, ColumnC, "NYSE:BA");
                var value = FastGetCellValue("Input", 9, ColumnC);
                if (value != "The Boeing Company")
                    return false;

                FastSetCellValue("Input", 5, ColumnC, "NYSE:A");
                value = FastGetCellValue("Input", 9, ColumnC);
                if (value != "Agilent Technologies, Inc.")
                    return false;

                return true;

            }, 10 * 60 * 1000);

            if (ok)
            {
                connectedSignal.Set();
            }
            else
            {
                failedSignal.Set();
                throw new CiqInactiveException();
            }
        }

        private void SetSpreadsheetInputs()
        {
            var settings = reportSettings;
            var fu = reportSettings.TimePeriodSettings;
            
            try
            {
                var commands = new StringBuilder();
                // AutomaticCalculation(false);

                // set the number of peers to the minimum to speed things up
                AppendUseWorksheet(commands, inputWorkSheet);
                AppendSetCellValue(commands, 16, ColumnC, 1);
                AppendSetCellValue(commands, 5, ColumnC, fu.TickerSymbol);
                if (!string.IsNullOrEmpty(fu.TimePeriodType))
                    AppendSetCellValue(commands, 5, ColumnI, (fu.TimePeriodType[0] == 'Q') ? "Quarter" : "Annual");

                // Track();
                AppendSetCellValue(commands, 9, ColumnH, fu.FirstPeriod);
                AppendSetCellValue(commands, 9, ColumnI, fu.LastPeriod);
                AppendSetCellValue(commands, 12, ColumnH, fu.DecompositionBegin);
                AppendSetCellValue(commands, 12, ColumnI, fu.DecompositionEnd);
                AppendSetCellValue(commands, 16, ColumnH, fu.PeerFirstPeriod);
                AppendSetCellValue(commands, 16, ColumnI, fu.PeerLastPeriod);

                // Track();
                if (!string.IsNullOrEmpty(settings.ReportType))
                {
                    var c = settings.ReportType[0];
                    if (c == 'F')
                        AppendSetCellValue(commands, 5, 20, "1");
                    else // if (c == 'N')
                        AppendSetCellValue(commands, 5, 20, "2");
                    // else
                    //    FastSetInputTextValue(inputWorksheet,5, 20, "3");
                }

                var response = SendCommandsToAddIn(commands);
                if (IsErrorResponse(response))
                    throw new CiqInactiveException();
                IncrementCurrentStep();

                int peerCount = 0;
                //FastSetCellValue(inputWorkSheet,  5, ColumnC, settings.TickerSymbol);
                //FastSetCellValue(inputWorkSheet, 19, ColumnE, settings.PeersShortNames[0]);

                if (settings.Peers != null)
                {
                    commands.Clear();
                    AppendUseWorksheet(commands, impDataWorkSheet);

                    for (var peerIndex = 0; peerIndex < settings.Peers.Length; ++peerIndex)
                    {
                        if (!string.IsNullOrWhiteSpace(settings.Peers[peerIndex]))
                        {
                            Log("Setting peer " + peerIndex + " to " + settings.Peers[peerIndex]);
                            AppendSetCellValue(commands, 4 + peerCount, ColumnL, settings.Peers[peerIndex]);
                            peerCount++;
                        }
                    }

                    while (peerCount < 10)
                    {
                        AppendSetCellValue(commands, 4 + peerCount, ColumnL, settings.Peers[peerCount]);
                       ++peerCount;
                    }
                    SendCommandsToAddIn(commands);

                    commands.Clear();
                    AppendUseWorksheet(commands, inputWorkSheet);
                    peerCount = 0;
                    for (var peerIndex = 0; peerIndex < settings.Peers.Length; ++peerIndex)
                    {
                        if (!string.IsNullOrWhiteSpace(settings.Peers[peerIndex]))
                        {
                            Log("Setting peer " + peerIndex + " to " + settings.Peers[peerIndex]);
                            AppendSetCellValue(commands, 20 + peerCount, ColumnE, settings.PeersShortNames[peerIndex]);
                            peerCount++;
                        }
                    }
                }

                SendCommandsToAddIn(commands);

                if (peerCount == 0)
                    FastSetCellValue(exportNWorkSheet, 1, ColumnF, 1);
                else
                    FastSetCellValue(exportNWorkSheet, 1, ColumnF, 2);

                IncrementCurrentStep();

                //FastSetCellValue(inputWorkSheet, 16, ColumnC, peerCount);
                IncrementCurrentStep();

                // Track();
            }
            catch (Exception ex)
            {
                fu.Error = ex.Message;
            }
            finally
            {
                // AutomaticCalculation(true);
            }
        }

        private void DoPeerGeneration()
        {
            report.Companies.Clear();
        }

        private void DoExcelGeneration()
        {
            try
            {
                // ExcelApp.DisplayAlerts = false;
                RunScript("UpdateAnalysis");
                // IncrementCurrentStep();
                ExcelApp.DisplayAlerts = false;
            }
            catch (Exception ex)
            {
                Log(ex);
                throw;
            }
            finally
            {
            }
        }

        #region Old PowerPoint Generation Code

        private void DoPowerPointGeneration()
        {
            this.Enter();

            try
            {
                var templateFilename = client.Server.MapInputFilename("Template.pptx");
                var templateCopyFilename = client.Server.MapOutputFilename(Path.GetFileNameWithoutExtension(WorkbookFilename) + ".pptx");
                Info("templateCopyFilename = " + templateCopyFilename);

                File.Copy(templateFilename, templateCopyFilename, true);
                var a = File.GetAttributes(templateCopyFilename);
                File.SetAttributes(templateCopyFilename, a & ~System.IO.FileAttributes.ReadOnly);
                IncrementCurrentStep();

                var finalName = reportSettings.TickerSymbol.Replace(':', '_') + "_" + DateTime.Now.ToString("yyyy_MM_dd");
                var tempIndex = 1;
                var probe = client.Server.MapOutputFilename(finalName + ".pptx");
                while (File.Exists(probe))
                {
                    ++tempIndex;
                    probe = client.Server.MapOutputFilename(finalName + "_" + tempIndex + ".pptx");
                }

                report.Filename = probe;
                var dir = Path.GetDirectoryName(templateCopyFilename);
                var tcfn = Path.GetFileName(templateCopyFilename);
                var rfn = Path.GetFileName(report.Filename);
                Info("CreatePowerPoint(" + dir + "," + tcfn + "," + rfn + ")");

                ExcelApp.DisplayAlerts = false;
                ExcelApp.Run("CreatePowerPoint", dir, tcfn, Path.GetFileName(report.Filename));
                ExcelApp.DisplayAlerts = false;
                IncrementCurrentStep();
            }
            catch (Exception ex)
            {
                Log(ex);
                throw;
            }
            finally
            {
                this.Exit();
            }
        }

        #endregion Old PowerPoint Generation Code

        #region New PowerPoint Generation Code

        private void DoBetterPowerPointGeneration()
        {
            Enter();

            var templateFilename = client.Server.MapInputFilename("Template.pptx");

            var baseName = reportSettings.TickerSymbol.Replace(':', '_') + "_" + DateTime.Now.ToString("yyyy_MM_dd");
            var tempIndex = 1;
            var workingFilename = client.Server.MapOutputFilename(baseName + ".pptx");
            while (File.Exists(workingFilename))
            {
                ++tempIndex;
                workingFilename = client.Server.MapOutputFilename(baseName + "_" + tempIndex + ".pptx");
            }

            File.Copy(templateFilename, workingFilename, true);
            var a = System.IO.File.GetAttributes(workingFilename);
            File.SetAttributes(workingFilename, a & ~System.IO.FileAttributes.ReadOnly);
            report.Filename = workingFilename;
            IncrementCurrentStep();

            Track();
            var presentations = powerpointApp.Presentations;
            PowerPoint.Presentation presentation = null;
            try
            {
                IncrementCurrentStep();
                presentations.Open(workingFilename, False, False, False);
                presentation = powerpointApp.Presentations[workingFilename];
                var originalSlideCount = presentation.Slides.Count;
                Track();

                // foreach (PowerPoint.Slide slide in presentation.Slides)
                //     Log("SLIDE " + slide.Name + " ID=" + slide.SlideID + " INDEX=" + slide.SlideIndex + " NUMBER=" + slide.SlideNumber);

                IncrementCurrentStep();
                try
                {
                    CreatePresentation(presentation);
                }
                catch (Exception ex)
                {
                    Log(ex);
                    throw;
                }

                IncrementCurrentStep();

                IncrementCurrentStep();
                for (var i = 0; i < originalSlideCount; ++i)
                    presentation.Slides[1].Delete();
                IncrementCurrentStep();
            }
            finally
            {
                if (presentation != null)
                {
                    presentation?.Save();
                    presentation?.Close();
                    ReleaseComObject(presentation);
                    presentation = null;
                }
                // presentations.Close();
                if (presentations != null)
                {
                    ReleaseComObject(presentations);
                    presentations = null;
                }
            }
        }


        private void CreatePresentation(PowerPoint.Presentation presentation)
        {
            var originalSlideCount = presentation.Slides.Count;
            var tmStart = DateTime.Now;
            var monthYear = tmStart.ToString("MMMM yyyy");

            // var overview = GetWorksheet(wkBook, "Overview");
            // var company = GetCellText(_input9, 3);
            var shortCompany = report.Companies[0].ShortName;
            // Track();
            // var impData = _wksheets["ImpData"];
            // var targetIndustry = GetRange(impData, "TargetIndustry").Value;

            IncrementCurrentStep();
            var profileType = FastGetCellValue(exportNWorkSheet, 1, 6);
            IncrementCurrentStep();

            if (reportSettings.ReportType[0] == 'F')
                exportWorkSheet = exportFWorkSheet;
            else // if (ReportSettings.ReportType[0] == 'N')
                exportWorkSheet = exportNWorkSheet;

            IncrementCurrentStep();
            float conversionFactor = 72F;
            //GetRange(exportN, "pMaxHeight").Value = presentation.PageSetup.SlideHeight * 0.75 / conversionFactor;
            //GetRange(exportN, "pMaxWidth").Value = presentation.PageSetup.SlideWidth * 0.95 / conversionFactor;
            presentation.PageSetup.FirstSlideNumber = 0;

            // Count the total number of slides
            //var startCell = GetRange(_exportWorkSheet, "eStart");
            var startRow = 6;
            var r = startRow;
            var totalSlides = 0;
            while (true)
            {
                Log("Scanning Row " + r);
                var value = FastGetCellValue(exportWorkSheet, r, ColumnC);
                if (string.Compare(value, "True", true) == 0)
                    ++totalSlides;
                ++r;

                value = FastGetCellValue(exportWorkSheet, r, ColumnB);
                if (value == "Exit")
                    break;
            };

            var slideNumber = 0;
            var documentFooter = shortCompany + " " +
                    ((profileType == "1") ? "Company Analysis" : "Company and Peer Analysis") +
                    " | " + monthYear;

            // var mainTitleTemplate = presentation.Slides[1];
            // var tocTemplate = presentation.Slides[2];
            var sectionTitleTemplate = presentation.Slides[3];
            var pageTemplate = presentation.Slides[4];
            // var appendixTitleTemplate = presentation.Slides[5];
            // var needMoreInfoTemplate = presentation.Slides[6];
            // var aboutTemplate = presentation.Slides[7];
            // var preparerTemplate = presentation.Slides[8];
            // var disclaimerTemplate = presentation.Slides[9];
            // var blankTemplate = presentation.Slides[10];
            PowerPoint.Slide toc = null;
            var mostRecentlyReportedSlideNumber = -1;

            r = startRow;
            do
            {
                Log("Running Row " + r);
                IncrementCurrentStep();

                //for (var i = 2; i < columns.Length; ++i)
                //    columns[i] = _FastGetCellText(exportWorkSheet,r, i);
                // Log("Row " + r);

                var values = FastGetCellValues(exportWorkSheet, r, 3, 7);
                var r3 = values[0];
                if (r3.Compare("true") != 0)
                {
                    // Log("Skipping row " + r);
                    r++;
                }
                else
                {
                    ++slideNumber;
                    string r2; // = GetCellText(_exportWorkSheet, r, 2);
                    // int srcSld = int.TryParse(Right(r2, 2), out var temp) ? temp : 0;
                    // presentation.Designs.Load(null); // ?
                    var r4 = values[1];
                    var r5 = values[2];
                    var r6 = values[3];
                    var r7 = values[4];
                    // Log("Processing row " + r + " r4=" + r4 + " r5=" + r5 + " r6=" + r6 + " r7=" + r7);

                    if (r4.Compare("PPT") == 0)
                    {
                        var originalSlide = int.Parse(Right(r5, 2));
                        var slide = DuplicateSlide(presentation, originalSlide);
                        Debug.Assert(slide != null);

                        if (originalSlide == 2)
                            toc = slide;

                        if (r6 == "AppendixSectionRoadMap")
                        {
                            if (toc != null)
                            {
                                var shape = toc.Shapes["Table Appendix"];
                                shape.Table.Rows[2].Cells[3].Shape.TextFrame.TextRange.Text = (presentation.Slides.Count - originalSlideCount - 1).ToString();
                                slide.Shapes[1].TextFrame.TextRange.Text = r7;
                            }
                        }

                        if (slideNumber == 1)
                        {
                            foreach (PowerPoint.Shape shape in slide.Shapes)
                            {
                                if (shape.Name == "Date")
                                {
                                    shape.TextFrame.TextRange.Text = monthYear;
                                }
                                else if (shape.Name == "Title")
                                {
                                    shape.TextFrame.TextRange.Text = shortCompany;
                                }
                                else if (shape.Name == "Subheading")
                                {
                                    if (profileType == "1")
                                        shape.TextFrame.TextRange.Text = "Company Analysis";
                                    else
                                        shape.TextFrame.TextRange.Text = "Company and Peer Analysis";
                                }
                            }

                            foreach (PowerPoint.Shape shape in presentation.SlideMaster.Shapes)
                            {
                                if (shape.Name == "DocumentFooter")
                                {
                                    shape.TextFrame.TextRange.Text = documentFooter;
                                    shape.TextFrame.TextRange.Font.Size = 9;
                                }
                            }
                        }
                        else
                        {
                            foreach (PowerPoint.Shape shape in slide.Design.SlideMaster.Shapes)
                            {
                                if (shape.Name == "DocumentFooter")
                                {
                                    shape.TextFrame.TextRange.Text = documentFooter;
                                    shape.TextFrame.TextRange.Font.Size = 9;
                                }
                            }
                        }

                        ++r;
                    }
                    else // if (r4.Compare("Excel") == 0)
                    {
                        try
                        {
                            Debug.Assert(r4 == "Excel");
                            PowerPoint.Slide slide; // = null;

                            if (r6 == "SectionRoadMap")
                            {
                                var shape = toc.Shapes["Table Contents"];

                                int i;
                                if (shape.Table.Rows[1].Cells[1].Shape.TextFrame.TextRange.Text == "1")
                                {
                                    shape.Table.Rows.Add();
                                    i = shape.Table.Rows.Count;
                                }
                                else
                                {
                                    i = 1;
                                }
                                shape.Table.Rows[i].Cells[1].Shape.TextFrame.TextRange.Text = i.ToString();
                                shape.Table.Rows[i].Cells[2].Shape.TextFrame.TextRange.Text = r7;
                                var pageNumber = slideNumber - 1;
                                shape.Table.Rows[i].Cells[3].Shape.TextFrame.TextRange.Text = pageNumber.ToString();

                                slide = DuplicateSlide(sectionTitleTemplate);
                            }
                            else
                            {
                                slide = DuplicateSlide(pageTemplate);
                            }

                            // Log("Duplicating source slide " + ppSld);

                            foreach (PowerPoint.Shape shape in slide.Design.SlideMaster.Shapes)
                            {
                                if (shape.Name == "DocumentFooter")
                                {
                                    shape.TextFrame.TextRange.Text = documentFooter;
                                    shape.TextFrame.TextRange.Font.Size = 9;
                                    break;
                                }
                            }
                            do
                            {
                                string sldText;
                                ShapeSpecification shpSpn;

                                if (r6 == "RoadMap")
                                {
                                    sldText = r7;
                                    shpSpn = RoadmapSpecification();
                                    var shpNum = InsertPlaceholder(slide, shpSpn);
                                    FormatShape(slide, shpNum, shpSpn, sldText);
                                }
                                else if (r6 == "SectionRoadMap")
                                {
                                    //Log("presentation.Slides.Count = " + presentation.Slides.Count);
                                    //Log("presentation.Slides[" + slideNumber + "].Shapes.Count = " + slide.Shapes.Count);
                                    //foreach (PowerPoint.Shape s in slide.Shapes)
                                    //{
                                    //    Log("Id=" + s.Id + " Name=" + s.Name + " Type=" + s.Type + " Alt=" + s.AlternativeText);
                                    //}
                                    var shape = FindShape(slide, 1);
                                    var tf = shape.TextFrame;
                                    var tr = tf.TextRange;
                                    tr.Text = r7;
                                }
                                else if (r6 == "AppendixSectionRoadMap")
                                {
                                    var shape = FindShape(slide, 1);
                                    shape.TextFrame.TextRange.Text = r7;
                                }
                                else if (r6 == "SlideTitle")
                                {
                                    sldText = r7;
                                    // Track();
                                    shpSpn = SlideTitleSpecification();
                                    // Track();
                                    var shpNum = InsertPlaceholder(slide, shpSpn);
                                    // Track();
                                    FormatShape(slide, shpNum, shpSpn, sldText);
                                    // Track();
                                }
                                else if (r6 == "NoteBox")
                                {
                                    sldText = r7;
                                    shpSpn = NoteBoxSpecification();
                                    var shpNum = InsertPlaceholder(slide, shpSpn);
                                    FormatShape(slide, shpNum, shpSpn, sldText);
                                }
                                else if (r6 == "SlideSource")
                                {
                                    sldText = r7;
                                    shpSpn = SourceBoxSpecification();
                                    var shpNum = InsertPlaceholder(slide, shpSpn);
                                    FormatShape(slide, shpNum, shpSpn, sldText);

                                }
                                else if (r6 == "TextBox")
                                {
                                    sldText = r7;
                                    var tText = sldText.Split(new[] { " vbNewLine " }, StringSplitOptions.RemoveEmptyEntries);
                                    sldText = string.Join("\n", tText);
                                    shpSpn = TextBoxSpecification();
                                    var shpNum = InsertPlaceholder(slide, shpSpn);
                                    FormatShape(slide, shpNum, shpSpn, sldText);
                                    // FormatBulletInShape(sPPT, newPPT, sldNum, shpNum); // TODO
                                }
                                else if ((r6 == "Picture") || (r6 == "Picture!"))
                                {
                                    //var r7cell = GetCell(_exportWorkSheet, r, 7);
                                    var wksName1 = r5;
                                    var wkSheet1 = wksName1;
                                    if (FindWorksheet(wkSheet1) != null)
                                    {
                                        try
                                        {
                                            if (CopyGraphsAsPictures || (r6 == "Picture!"))
                                            {
                                                FindWorksheet(wkSheet1).Activate();

                                                //Yield(1000);
                                                var picture = GetShape(wkSheet1, r7);
                                                Info("Copying picture " + r7);
                                                if (picture != null)
                                                {
                                                    picture.Copy();
                                                    //Yield(1000);
                                                    var shape = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteShape);
                                                    //.AddShape(Office.MsoAutoShapeType..msoShapeRectangle /*templatePicture.Type*/, picture.Left, picture.Top, picture.Width, picture.Height); //  'ppPasteShape) ' ppPasteOLEObject) ' ppPasteDefault) */

                                                    var sides = FastGetCellValues(exportWorkSheet, r, 8, 11);
                                                    var left = float.Parse(sides[0]);
                                                    var top = float.Parse(sides[1]);
                                                    var width = float.Parse(sides[2]);
                                                    var height = float.Parse(sides[3]);

                                                    shape.Left = left * conversionFactor;
                                                    shape.Top = top * conversionFactor;
                                                    if (width > 0)
                                                        shape.Width = width * conversionFactor;
                                                    if (height > 0)
                                                        shape.Height = height * conversionFactor;
                                                }
                                            }
                                            else if (r7 == "-2146826246")
                                            {
                                                Warning("#NA for row " + r + ", slide " + slideNumber);
                                            }
                                            else
                                            {
                                                FindWorksheet(wkSheet1).Activate();
                                                //Yield(1000);
                                                var excelCharts = (Excel.ChartObjects)(FindWorksheet(wkSheet1).ChartObjects());
                                                this.Trace("Charts on worksheet '" + wkSheet1 + "'");
                                                for (var i = 1; i <= excelCharts.Count; ++i)
                                                {
                                                    string s = excelCharts.Item(i).Name.ToString();
                                                    this.Trace("    " + i + ") " + s);
                                                }
                                                this.Trace("Looking for '" + r7 + "'");

                                                Excel.ChartObject excelChart = null;
                                                try
                                                {
                                                    excelChart = excelCharts.Item(r7) as Excel.ChartObject;
                                                }
                                                catch
                                                {
                                                    excelChart = null;
                                                }

                                                if (excelChart != null)
                                                {
                                                    Info("Copying chart " + r7);
                                                    excelChart.Copy();
                                                    //Yield(1000);
                                                    var shapeRange = slide.Shapes.PasteSpecial(DataType: PowerPoint.PpPasteDataType.ppPasteDefault, Link: Microsoft.Office.Core.MsoTriState.msoFalse);
                                                    var shape = slide.Shapes[slide.Shapes.Count - 1];

                                                    var sides = FastGetCellValues(exportWorkSheet, r, 8, 11);
                                                    var left = float.Parse(sides[0]);
                                                    var top = float.Parse(sides[1]);
                                                    var width = float.Parse(sides[2]);
                                                    var height = float.Parse(sides[3]);

                                                    shapeRange.Left = left * conversionFactor;
                                                    shapeRange.Top = top * conversionFactor;
                                                    if (width > 0)
                                                        shapeRange.Width = width * conversionFactor;
                                                    if (height > 0)
                                                        shapeRange.Height = height * conversionFactor;

                                                    /*
                                                    Log("Trying to break links");
                                                    shape.Chart.ChartData.BreakLink();
                                                    Log("Links broken");
                                                    */

                                                    /*
                                                    var data = new List<List<object>>();
                                                    var excelSeriesCollection = excelChart.Chart.SeriesCollection();
                                                    foreach (var excelSeries in excelSeriesCollection)
                                                    {
                                                        var list = new List<object>();
                                                        data.Add(list);
                                                        var s = (Excel.Series) excelSeries;
                                                        var t = s.Values;
                                                        var w = (object) t;
                                                        var v = (Array) w;
                                                        foreach (var u in v)
                                                        {
                                                            // Log(s.Name + " value " + u);
                                                            list.Add(u);
                                                        }
                                                    }
                                                    */

                                                    /*
                                                    var powerPointChart = (PowerPoint.Shape) slide.Shapes[slide.Shapes.Count - 1];
                                                    */

                                                    /*
                                                    var dataSheet = powerPointChart.Application.DataSheet;

                                                    var rows = data.Count;
                                                    var columns = data[0].Count;
                                                    int row = 0;
                                                    foreach (var list in data)
                                                    {
                                                        int col = 0;
                                                        foreach (var obj in list)
                                                        {
                                                            dataSheet.Cell(row, col++).Shape.TextFrame2.TextRange.Text = obj.ToString();
                                                        }
                                                        ++row;
                                                    }
                                                    */

                                                    Info("Copied chart " + r7 + " to " + left + ", " + top + ", " + width + ", " + height);
                                                    ZapComObject(ref excelChart);
                                                }
                                                else
                                                {
                                                    Info("Charts on worksheet '" + wkSheet1 + "'");
                                                    for (var i = 1; i <= excelCharts.Count; ++i)
                                                    {
                                                        string s = excelCharts.Item(i).Name.ToString();
                                                        Info("    " + i + ") " + s);
                                                    }
                                                    Log("Chart not found: " + r7 + " from row " + r + ", slide " + slideNumber + ", worksheet " + wksName1);
                                                }
                                                ZapComObject(ref excelCharts);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            // Bad chart
                                            Log("Exception row " + r + ", slide " + slideNumber);
                                            Log(ex);
                                        }
                                        finally
                                        {
                                            wkSheet1 = null;
                                        }
                                    }
                                }
                                else if (r6 == "Range")
                                {
                                    Info("Resolving range on row " + r);
                                    var wksName1 = r5;
                                    var wkSheet1 = wksName1;
                                    if (wkSheet1 != null)
                                    {
                                        //var r7cell = exportWorkSheet.GetCell(r, 7);
                                        string cpyRng = FastGetCellFormula(exportWorkSheet, r, 7); // r7cell.Formula;
                                        Info("Formula is " + cpyRng);
                                        if (cpyRng.StartsWith("="))
                                            cpyRng = cpyRng.Substring(1);
                                        if (cpyRng.StartsWith("@"))
                                            cpyRng = cpyRng.Substring(1);
                                        if (cpyRng.StartsWith("+"))
                                            cpyRng = cpyRng.Substring(1);
                                        Info("Cleaned formula is " + cpyRng);

                                        if (!cpyRng.Contains(":"))
                                        {
                                            Info("Range " + cpyRng + " does not contain a colon");
                                            if (ParseCellAddress(cpyRng, out var wksName2, out var row, out var column))
                                            {
                                                string wkSheet2;
                                                if ((wksName2 == null) || (wksName2 == wksName1))
                                                {
                                                    wkSheet2 = wkSheet1;
                                                }
                                                else
                                                {
                                                    wkSheet2 = wksName2;
                                                }
                                                var redirect = FastGetCellValue(wkSheet2, row, column);
                                                Info("Redirected  " + cpyRng + " to " + redirect);
                                                cpyRng = redirect;
                                            }
                                            else
                                            {
                                                Log("Redirection failed for " + cpyRng);
                                            }
                                        }

                                        // Track();
                                        Yield(1000);
                                        Exception ex = null;
                                        bool isRange = !cpyRng.StartsWith("IF(");
                                        if (isRange)
                                        {
                                            try
                                            {
                                                if (cpyRng.StartsWith(wksName1 + "!") || cpyRng.StartsWith("'" + wksName1 + "'!"))
                                                {
                                                    int e = cpyRng.IndexOf('!');
                                                    cpyRng = cpyRng.Substring(e + 1);
                                                }
                                                Info("Attempting to copy from range " + cpyRng);
                                                FindWorksheet(wkSheet1).Range[cpyRng].Copy();
                                                Info("Successfully copied from range " + cpyRng);
                                            }
                                            catch (Exception x1)
                                            {
                                                Info("Failed to copy from range " + cpyRng + " from row " + r + ", slide " + slideNumber);
                                                isRange = false;
                                                ex = x1;
                                            }
                                        }

                                        if (!isRange)
                                        {
                                            // Track();
                                            try
                                            {
                                                Info("Attempting to copy from range given by " + cpyRng);
                                                cpyRng = FastGetCellValue(exportWorkSheet, r, 7);
                                                Info("Converted to " + cpyRng);
                                                FindWorksheet(wkSheet1).Range[cpyRng].Copy();
                                                Info("Successfully copied from range " + cpyRng);
                                                if (ex != null)
                                                {
                                                    Info("Exception handled");
                                                    ex = null;
                                                }
                                            }
                                            catch (Exception x2)
                                            {
                                                Log("Failed to copy from range " + cpyRng + " from row " + r + ", slide " + slideNumber);
                                                ex = x2;
                                            }
                                        }

                                        wkSheet1 = null;
                                        if (ex != null)
                                            throw ex;

                                        // Track();
                                        var sides = FastGetCellValues(exportWorkSheet, r, 8, 11);
                                        var left = float.Parse(sides[0]);
                                        var top = float.Parse(sides[1]);
                                        var width = float.Parse(sides[2]);
                                        var height = float.Parse(sides[3]);

                                        // Track();
                                        Yield(1000);
                                        var tmpShp = slide.Shapes.PasteSpecial(DataType: PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile);

                                        // Application.CutCopyMode = false;
                                        tmpShp.LockAspectRatio = True;

                                        tmpShp.Left = left * conversionFactor;
                                        tmpShp.Top = top * conversionFactor;

                                        // Track();
                                        if (width > 0)
                                            tmpShp.Width = width * conversionFactor;

                                        if (height > 0)
                                            tmpShp.Height = height * conversionFactor;

                                        // Track();
                                        if (tmpShp.Top + tmpShp.Height > 553.0359)
                                        {
                                            tmpShp.LockAspectRatio = True;
                                            tmpShp.Height = 553.0359F - tmpShp.Top;
                                        }

                                        // Track();
                                        if (tmpShp.Width > 9.87 * 72)
                                        {
                                            tmpShp.LockAspectRatio = True;
                                            tmpShp.Width = 9.87F * 72;
                                        }

                                        // Track();
                                        if (string.Compare(FastGetCellValue(exportWorkSheet,r, 12), "true", true) == 0)
                                            tmpShp.Left = (presentation.PageSetup.SlideWidth / 2F) - (tmpShp.Width / 2F);
                                        else
                                            tmpShp.Left = left * conversionFactor;
                                    }
                                }

                                ++r;

                                //for (var i = 2; i < columns.Length; ++i)
                                //    columns[i] = _FastGetCellText(exportWorkSheet,r, i);
                                //Log("ROW " + r + " |" + string.Join("|", columns, 2, columns.Length - 2));

                                r2 = FastGetCellValue(exportWorkSheet, r, 2);
                                if (r2 == "")
                                {
                                    // r4 = _FastGetCellText(exportWorkSheet,r, 4);
                                    // r5 = _FastGetCellText(exportWorkSheet,r, 5);
                                    r6 = FastGetCellValue(exportWorkSheet, r, 6);
                                    r7 = FastGetCellValue(exportWorkSheet, r, 7);
                                }

                            } while (r2 == "");
                        }
                        catch (Exception ex)
                        {
                            Log("Exception while handling row " + r + ", slide " + slideNumber);
                            Log(ex);
                        }
                    }
                }

                if (mostRecentlyReportedSlideNumber < slideNumber)
                {
                    mostRecentlyReportedSlideNumber = slideNumber;
                }

            } while (FastGetCellValue(exportWorkSheet, r, 2) != "Exit");

            DeleteLastSlide(presentation); // I don't know why we need to do this, but we do

            FindWorksheet(inputWorkSheet).Activate();
        }

        private bool ParseCellAddress(string address, out string worksheetName, out int row, out int column)
        {
            var excl = address.IndexOf('!');
            if (excl < 0)
            {
                worksheetName = null;
            }
            else
            {
                worksheetName = address.Substring(0, excl);
                if (worksheetName.StartsWith("'"))
                    worksheetName = worksheetName.Substring(1, worksheetName.Length - 2);
                address = address.Substring(excl + 1);
            }


            row = 0;
            column = 0;
            foreach (var c in address)
            {
                if (c == '$')
                {
                    // do nothing
                }
                else if (char.IsDigit(c))
                {
                    row *= 10;
                    row += (c - '0');
                }
                else
                {
                    column *= 26;
                    column += c - 'A';
                }
            }
            column++;
            return true;
        }

        #endregion New PowerPoint Generation Code

        private string Right(string input, int count) => input.Substring(input.Length - count);
        private string Right(Excel.Range range, int count) => Right(range.Value.ToString(), count);
        private int RGB(int r, int g, int b) => (int)(r | (g << 8) | (b << 16));

        #region Shape Specifications

        private void FormatShape(PowerPoint.Slide slide, int shpNum, ShapeSpecification ss, string sldText = "")
            => FormatShape(slide, shpNum, ss, true, true, sldText);

        private void FormatShape(PowerPoint.Slide slide, int shpNum, ShapeSpecification ss, bool textAdjust, bool sizeAdjust, string sldText)
        {
            // Enter();

            // Log("shpNum = " + shpNum);
            var shape = slide.Shapes[shpNum];
            var shpSpn = ss.Raw;
            shape.Name = shpSpn[0];

            // Track();
            shape.Line.Visible = shpSpn[1];
            if (shape.Line.Visible == True)
            {
                if (shpSpn[2] != null)
                    shape.Line.ForeColor.RGB = shpSpn[2];
                if (shpSpn[3] != null)
                    shape.Line.BackColor.RGB = shpSpn[3];
                if (shpSpn[4] != null)
                    shape.Line.Weight = shpSpn[4];
                if (shpSpn[5] != null)
                    shape.Line.Style = shpSpn[5];
            }

            // Track();
            shape.Fill.Visible = shpSpn[11];
            if (shape.Fill.Visible == True)
            {
                if (shpSpn[12] != null)
                    shape.Fill.ForeColor.RGB = shpSpn[12];
                if (shpSpn[13] != null)
                    shape.Fill.Transparency = shpSpn[13];
            }

            // Track();
            shape.Shadow.Visible = shpSpn[14];
            if (shape.Shadow.Visible == True)
            {
                if (shpSpn[15] != null)
                    shape.Shadow.ForeColor.RGB = shpSpn[15];
                if (shpSpn[16] != null)
                    shape.Shadow.Blur = shpSpn[16];
                if (shpSpn[17] != null)
                    shape.Shadow.Size = shpSpn[17];
                if (shpSpn[18] != null)
                    shape.Shadow.Transparency = shpSpn[18];
                if (shpSpn[19] != null)
                    shape.Shadow.OffsetX = shpSpn[19];
                if (shpSpn[20] != null)
                    shape.Shadow.OffsetY = shpSpn[20];
            }

            // Track();
            if (sizeAdjust)
            {
                shape.Left = (float)shpSpn[21];
                shape.Top = (float)shpSpn[22];
                shape.Width = (float)shpSpn[23];
                shape.Height = (float)shpSpn[24];
            }

            // Track();
            shape.Rotation = shpSpn[25];
            shape.LockAspectRatio = shpSpn[26];

            // Track();
            var textFrame2 = shape.TextFrame2;
            textFrame2.AutoSize = ss.TextFrameAutoSize;
            textFrame2.Orientation = ss.TextFrameOrientation;
            textFrame2.VerticalAnchor = ss.TextFrameVerticalAnchor;
            textFrame2.MarginLeft = ss.TextFrameMarginLeft;
            textFrame2.MarginRight = ss.TextFrameMarginRight;
            textFrame2.MarginTop = ss.TextFrameMarginTop;
            textFrame2.MarginBottom = ss.TextFrameMarginBottom;
            textFrame2.WordWrap = ss.TextFrameWordWrap;

            // Track();
            var textFrame = shape.TextFrame;
            var textRange = textFrame.TextRange;
            var textRange2 = textFrame2.TextRange;
            if (textAdjust)
                textRange.Text = sldText?.TrimEnd() ?? "";
            if (shpSpn[42] != null)
                textRange.Font.Name = shpSpn[42].ToString();
            if (shpSpn[43] != null)
                textRange.Font.Bold = (Office.MsoTriState)shpSpn[43];
            if (shpSpn[44] != null)
                textRange.Font.Italic = (Office.MsoTriState)shpSpn[44];
            if (shpSpn[45] != null)
                textRange.Font.Underline = shpSpn[45];
            if (shpSpn[46] != null)
                textRange.Font.Size = (float)shpSpn[46];
            if (shpSpn[47] != null)
                textRange2.Font.Fill.ForeColor.RGB = (int)shpSpn[47];
            if (shpSpn[48] != null)
                textRange.Font.Shadow = (Office.MsoTriState)shpSpn[48];
            //  textRange.Paragraphs().ParagraphFormat.Bullet = shpSpn(51)

            // Track();
            var paragraphFormat = textRange.ParagraphFormat;
            if (shpSpn[52] != null)
                paragraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)shpSpn[52];
            if (shpSpn[53] != null)
                paragraphFormat.HangingPunctuation = (Office.MsoTriState)shpSpn[53];
            if (shpSpn[54] != null)
                paragraphFormat.SpaceBefore = (float)shpSpn[54];
            if (shpSpn[55] != null)
                paragraphFormat.SpaceAfter = (float)shpSpn[55];
            if (shpSpn[56] != null)
                paragraphFormat.SpaceWithin = (float)shpSpn[56];
            if (shpSpn[57] != null)
                paragraphFormat.Parent.Parent.Ruler.Levels(1).FirstMargin = shpSpn[57];
            if (shpSpn[58] != null)
                paragraphFormat.Parent.Parent.Ruler.Levels(1).LeftMargin = shpSpn[58];

            // Exit();
        }



        private ShapeSpecification RoadmapSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Roadmap"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 39.99; //0                        ; //Shape Left
            shpSpecifications[22] = 91.84; //8.625039                 ; //Shape Top
            shpSpecifications[23] = 713.24; //792                     ; //Shape Width
            shpSpecifications[24] = 21.81; //30.25                    ; //Shape Height
            shpSpecifications[25] = 0; //Rotation
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;   // Orientation
            // ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;                 // Vertical Anchor
            ss.TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            // ss.TextFrameMarginLeft = 0;                                                         // Margin Left
            // ss.TextFrameMarginRight = 0;                                                        // Margin Right
            // ss.TextFrameMarginTop = 0;                                                          // Margin Top
            // ss.TextFrameMarginBottom = 0;                                                       // Margin Bottom
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                  // Word Wrap

            shpSpecifications[41] = "ROADMAP"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = False; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 18; //Font Size
            shpSpecifications[47] = RGB(94, 138, 180); //(0, 43, 73)  ; //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence                  ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                  ; //Paragraph Alignment
            shpSpecifications[53] = False; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 0; //Ruler Level 1 Left Margin

            return ss;
        }


        private ShapeSpecification SlideTitleSpecification()
        {
            // Enter();

            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            // Track();
            shpSpecifications[0] = "Slide Title"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = RGB(94, 138, 180); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 38.625; //34.53 ; //17.24992         ; //Shape Left
            shpSpecifications[22] = 15.02; //55.91992 ; //39.74984       ; //Shape Top
            shpSpecifications[23] = 714.61; //684.75                  ; //Shape Width
            shpSpecifications[24] = 64.638; //28.08008 ; //42.33023      ; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;   // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;                 // Vertical Anchor
            ss.TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeNone; // Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            // ss.TextFrameMarginLeft = 0;                                                         // Margin Left
            // ss.TextFrameMarginRight = 0;                                                        // Margin Right
            // ss.TextFrameMarginTop = 0;                                                          // Margin Top
            ss.TextFrameMarginBottom = 3.6F;                                                       // Margin Bottom
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                  // Word Wrap

            shpSpecifications[41] = "Slide Title"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = False; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 28; //Font Size
            shpSpecifications[47] = RGB(94, 138, 180); //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence               ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                  ; //Paragraph Alignment
            shpSpecifications[53] = False; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 0; //Ruler Level 1 Left Margin

            // Exit();
            return ss;
        }


        private ShapeSpecification NoteBoxSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Note Box"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 32.7874; //Shape Left
            shpSpecifications[22] = 554.5049; //Shape Top
            shpSpecifications[23] = 739.4126; //675.2126              ; //Shape Width
            shpSpecifications[24] = 26.65779; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;   // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;                 // Vertical Anchor
            // ss.TextFrameAutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;                        // Auto Size
            ss.TextFrameMarginLeft = 7.2F;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                  // Word Wrap

            shpSpecifications[31] = Office.MsoTextOrientation.msoTextOrientationHorizontal; //Orientation
            shpSpecifications[32] = Office.MsoVerticalAnchor.msoAnchorBottom; //Vertical Anchor
            shpSpecifications[33] = 0; //ppAutoSizeNone                  ; //Auto Size
            shpSpecifications[34] = 7.2; //Margin Left
            shpSpecifications[35] = 7.2; //Margin Right
            shpSpecifications[36] = 3.6; //Margin Top
            shpSpecifications[37] = 3.6; //Margin Bottom
            shpSpecifications[38] = True; //Word Wrap

            shpSpecifications[41] = "Note:\t1) Sample Notes\n\t2) Sample Notes"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = False; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 9; //Font Size
            shpSpecifications[47] = RGB(118, 113, 113); //(0, 53, 95) ; //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence                  ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                     ; //Paragraph Alignment
            shpSpecifications[53] = True; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 32; //Ruler Level 1 Left Margin

            return ss;
        }


        private ShapeSpecification SourceBoxSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Source Box"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 32.7874; //Shape Left
            shpSpecifications[22] = 553.0359; //Shape Top
            shpSpecifications[23] = 675.2126; //Shape Width
            shpSpecifications[24] = 16.96409; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;    // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;                  // Vertical Anchor
            // ss.TextFrameAutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;                         //Auto Size
            ss.TextFrameMarginLeft = 7.2F;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                   //Word Wrap

            shpSpecifications[41] = "Source:\tSample Sources; (continued in slide note)"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = False; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 9; //Font Size
            shpSpecifications[47] = RGB(118, 113, 113); //(0, 53, 95) ; //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence                  ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                     ; //Paragraph Alignment
            shpSpecifications[53] = False; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 32; //Ruler Level 1 Left Margin

            return ss;
        }


        private ShapeSpecification ChartTitleSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Chart Title"; //Name"

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = True; //Fill Visible
            shpSpecifications[12] = null; // RGB(0, 43, 73); //Fill Color
            shpSpecifications[13] = 0; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 366; //Shape Left
            shpSpecifications[22] = 288; //Shape Top
            shpSpecifications[23] = 318; //Shape Width
            shpSpecifications[24] = 36; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;    // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;                  // Vertical Anchor
            // ss.TextFrameAutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;                         //Auto Size
            ss.TextFrameMarginLeft = 7.2F;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                   //Word Wrap

            shpSpecifications[41] = "Chart Title"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = True; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 12; //Font Size
            shpSpecifications[47] = RGB(255, 255, 255); //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence                  ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 2; //ppAlignCEnter                   ; //Paragraph Alignment
            shpSpecifications[53] = True; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 0; //Ruler Level 1 Left Margin


            return ss;
        }


        private ShapeSpecification DraftBoxSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Draft Box"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = null; //  RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 11.33858; //Shape Left
            shpSpecifications[22] = 11.33858; //Shape Top
            shpSpecifications[23] = 127.5591; //Shape Width
            shpSpecifications[24] = 30.33071; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;    // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;                  // Vertical Anchor
            ss.TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            ss.TextFrameMarginLeft = 7.2F;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                   //Word Wrap

            shpSpecifications[41] = "DRAFT"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = True; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 24; //Font Size
            shpSpecifications[47] = RGB(255, 0, 0); //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 3; //ppCaseUpper                     ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                     ; //Paragraph Alignment
            shpSpecifications[53] = False; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 0; //Ruler Level 1 Left Margin

            return ss;
        }


        private ShapeSpecification ConfidentialBoxSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Confidential Box"; //Name

            shpSpecifications[1] = False; //Line Visible

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = null; // RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 8.464882; //Shape Left
            shpSpecifications[22] = 512.4622; //Shape Top
            shpSpecifications[23] = 228.2018; //Shape Width
            shpSpecifications[24] = 20.59921; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;    // Orientation
            ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;                  // Vertical Anchor
            ss.TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            ss.TextFrameMarginLeft = 7.2F;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                   //Word Wrap

            shpSpecifications[41] = "A&M Confidential, not for distribution"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = True; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 11; //Font Size
            shpSpecifications[47] = RGB(255, 0, 0); //Font Color
            shpSpecifications[48] = False; //Shadow
            shpSpecifications[49] = 1; //ppCaseSentence                  ; //Case

            shpSpecifications[51] = False; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                     ; //Paragraph Alignment
            shpSpecifications[53] = False; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 0; //Ruler Level 1 Left Margin


            return ss;
        }


        private ShapeSpecification TextBoxSpecification()
        {
            var ss = new ShapeSpecification();
            var shpSpecifications = ss.Raw;

            shpSpecifications[0] = "Text Box"; //Name

            shpSpecifications[1] = False; //Line Visible
            shpSpecifications[2] = null; // RGB(0, 43, 73); //Line Fore Color
            shpSpecifications[3] = null; // RGB(0, 43, 73); //Line Back Color
            shpSpecifications[4] = 1; //Line Weight
            shpSpecifications[5] = Office.MsoLineStyle.msoLineSingle; //Line Style

            shpSpecifications[11] = False; //Fill Visible
            shpSpecifications[12] = RGB(255, 255, 255); //Fill Color
            shpSpecifications[13] = 1; //Fill Transparency
            shpSpecifications[14] = False; //Shadow Visible

            shpSpecifications[21] = 17.24992; //Shape Left
            shpSpecifications[22] = 87.88945; //Shape Top
            shpSpecifications[23] = 684.8504; //Shape Width
            shpSpecifications[24] = 373.3228; //Shape Height
            shpSpecifications[25] = 0; //Rotaion
            shpSpecifications[26] = False; //Lock Aspect Ratio

            // ss.TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;    // Orientation
            // ss.TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;               // Vertical Anchor
            ss.TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;
            ss.TextFrameMarginRight = 7.2F;
            ss.TextFrameMarginTop = 3.6F;
            ss.TextFrameMarginBottom = 3.6F;
            // ss.TextFrameWordWrap = Office.MsoTriState.msoTrue;                                   //Word Wrap

            shpSpecifications[41] = "Sample Text 1\nSample Text 2\nSample Text 3"; //Default Text
            shpSpecifications[42] = "Arial"; //Font Name
            shpSpecifications[43] = False; //Bold
            shpSpecifications[44] = False; //Italics
            shpSpecifications[45] = False; //Underline
            shpSpecifications[46] = 12; //Font Size
            shpSpecifications[47] = RGB(0, 43, 73); //Font Color
            shpSpecifications[48] = False; //Shadow
            ; //shpSpecifications[49] = ppCaseUpper                    ; //Case

            shpSpecifications[51] = True; //Paragraph Bullet
            shpSpecifications[52] = 1; //ppAlignLeft                     ; //Paragraph Alignment
            shpSpecifications[53] = True; //Paragraph Hanging Punctuation
            shpSpecifications[54] = 0; //Paragraph Space Before
            shpSpecifications[55] = 0; //Paragraph Space After
            shpSpecifications[56] = 1; //Paragraph Space Within
            shpSpecifications[57] = 0; //Ruler Level 1 First Margin
            shpSpecifications[58] = 27; //Ruler Level 1 Left Margin

            return ss;
        }

        #endregion

        internal PowerPoint.Slide DuplicateSlide(PowerPoint.Presentation presentation, int index)
        {
            Info("Duplicating slide index " + index + " (1)");
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (slide.SlideIndex == index)
                {
                    var generatedSlide = slide.Duplicate();
                    generatedSlide.MoveTo(presentation.Slides.Count);
                    return presentation.Slides[presentation.Slides.Count];
                }
            }
            return null;
        }

        internal void DeleteLastSlide(PowerPoint.Presentation presentation)
        {
            DeleteSlide(presentation, presentation.Slides.Count);
        }

        internal void DeleteSlide(PowerPoint.Presentation presentation, int index)
        {
            Info("Deleting slide index " + index);
            foreach (PowerPoint.Slide slide in presentation.Slides)
            {
                if (slide.SlideIndex == index)
                {
                    slide.Delete();
                    return;
                }
            }
            Log("Slide index " + index + " not found");
        }

        internal PowerPoint.Slide DuplicateSlide(PowerPoint.Slide slide)
        {
            Info("Duplicating slide index " + slide.SlideIndex + " (2)");
            var generatedSlide = slide.Duplicate();
            var presentation = (PowerPoint.Presentation)slide.Parent;
            generatedSlide.MoveTo(presentation.Slides.Count);
            return presentation.Slides[presentation.Slides.Count];
        }

        internal static PowerPoint.Shape FindShape(PowerPoint.Slide slide, int index)
        {
            return slide.Shapes[index];
        }

        internal Excel.Shape GetShape(string sheet, string name)
        {
            Enter();
            Trace("Looking for shape " + name);
            try
            {
                var shapes = FindWorksheet(sheet).Shapes;
                foreach (Excel.Shape shape in shapes)
                {
                    if (shape.Name == name)
                    {
                        Trace("Found shape " + name + " " + shape.GetType());
                        this.Exit();
                        return shape;
                    }
                }
                Warning("Failed to find shape " + name);
            }
            catch (Exception ex)
            {
                Log(ex);
            }
            Exit();
            return null;
        }

        internal int InsertPlaceholder(PowerPoint.Slide slide, ShapeSpecification ss)
        {
            var shapes = slide.Shapes;
            var oPlaceholder = shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, (float)ss.Raw[21], (float)ss.Raw[22], (float)ss.Raw[23], (float)ss.Raw[24]);
            oPlaceholder.TextFrame.TextRange.Text = "Sample";
            var shpName = oPlaceholder.Name;

            var shpNum = -1;
            for (var i = 1; i <= slide.Shapes.Count; ++i)
            {
                if (slide.Shapes[i].Name == shpName)
                {
                    shpNum = i;
                    break;
                }
            }

            return shpNum;
        }


        private void UnloadOfficeApps()
        {
            Enter();

            RunScript("ClearClipboard");

            ZapComObject(ref inputWorkSheet);
            ZapComObject(ref impDataWorkSheet);
            ZapComObject(ref exportNWorkSheet);
            ZapComObject(ref exportFWorkSheet);
            ZapComObject(ref exportWorkSheet);

            if (powerpointApp != null)
            {
                try { powerpointApp.Quit(); } catch { }
                for (var i = 0; i < 100; ++i)
                {
                    if (Marshal.FinalReleaseComObject(powerpointApp) > 0)
                    {
                        //Yield;
                    }
                }
                powerpointApp = null;

                // Console.WriteLine("UOA: 205");
            }

            // try { File.Delete(excelWorkingFilename); } catch { }

            // Console.WriteLine("UOA: Exit");
            Exit();
        }

        #region Dispose

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    Stop();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposed = true;
            }
            base.Dispose(disposing);
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~Server()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        #endregion
    }



}
