using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

namespace Acp
{
    public enum LogMessageType
    {
        Emergency = 0,  //	Emergency System is unusable
        Alert = 1,      //	Alert Action must be taken immediately
        Critical = 2,    //	Critical Critical conditions
        Error = 3,      //	Error Error conditions
        Warning = 4,    //	Warning Warning conditions
        Notice = 5,     //	Notice Normal but significant condition
        Info = 6,       //	Informational Informational messages
        Debug = 7,      //	Debug Debug-level messages

        Trace = Debug,
    }

    public struct LogSettings
    {
        public static readonly LogSettings All = new LogSettings(true);
        public static readonly LogSettings None = new LogSettings(true);
        private byte settings;
        private const byte toConsoleFlag = 0x01;
        private const byte toDebugFlag = 0x02;
        private const byte toRecorderFlag = 0x02;

        public bool ToConsole
        {
            get => (settings & toConsoleFlag) != 0;
            set
            {
                var current = (settings & toConsoleFlag) != 0;
                if (value != current)
                    settings ^= toConsoleFlag;
            }
        }

        public bool ToDebug
        {
            get => (settings & toDebugFlag) != 0;
            set
            {
                var current = (settings & toDebugFlag) != 0;
                if (value != current)
                    settings ^= toDebugFlag;
            }
        }

        public bool ToRecorder
        {
            get => (settings & toRecorderFlag) != 0;
            set
            {
                var current = (settings & toRecorderFlag) != 0;
                if (value != current)
                    settings ^= toRecorderFlag;
            }
        }

        public LogSettings(bool all) : this(all, all, all)
        {
        }
        public LogSettings(bool console, bool debug, bool recorder)
        {
            settings = 0;
            if (console)
                settings |= toConsoleFlag;
            if (debug)
                settings |= toDebugFlag;
            if (recorder)
                settings |= toRecorderFlag;
        }

        public LogSettings(LogSettings other)
        {
            settings = other.settings;
        }

        public bool NoLogging => (settings == 0);
    }

    public class LogSettingsCollection
    {
        private static readonly LogSettings[] defaultSettings
            = new LogSettings[] { /* Emergency */ new LogSettings(true),
                                  /* Alert */     new LogSettings(true),
                                  /* Critical*/   new LogSettings(true),
                                  /* Error */     new LogSettings(true),
                                  /* Warning */   new LogSettings(true),
                                  /* Notice */    new LogSettings(true),
                                  /* Info */      new LogSettings(true),
                                  /* Debug */     new LogSettings(false) };


        private readonly LogSettings[] settings;

        public LogSettingsCollection(LogSettingsCollection other = null)
        {
            settings = new LogSettings[defaultSettings.Length];
            var a = other?.settings ?? defaultSettings;
            for (var i = 0; i < settings.Length; ++i)
                settings[i] = new LogSettings(a[i]);
        }

        public LogSettings this[LogMessageType type]
        {
            get => settings[(int)type];
        }

        public void SetLevel(int level)
        {
            for (var i = 0; i < settings.Length; ++i)
            {
                settings[i] = new LogSettings(i <= level);
            }
        }

        public void SetLevel(LogMessageType level)
            => SetLevel((int)level);

        public void SetLevel(char c)
        {
            for (var i = 0; i < LogEntry.LogMessageTypeName.Length; ++i)
            {
                if (LogEntry.LogMessageTypeName[i] == c)
                {
                    SetLevel(i);
                    return;
                }
            }
        }
    }

    #region Log Entry

    public class LogEntry
    {
        public LogMessageType Type;
        public string Message;
        public string CallerMemberName;
        public string CallerFilePath;
        public int CallerLineNumber;
        public DateTime Time;
        public int ThreadId;
        public string ThreadName;
        public int ProcessId;
        public string AppName;
        private string fullMessage;
        public static readonly char[] LogMessageTypeName;

        static LogEntry()
        {
            LogMessageTypeName = new char[] { 'X', 'A', 'C', 'E', 'W', 'N', 'I', 'D' };
        }

        public LogEntry(LogMessageType type, long ticks, int procId, string procName, int threadId, string threadName, string message, string callerMemberName, string callerFilePath, int callerLineNumber)
        {
            Type = type;
            Message = message ?? "*";
            CallerMemberName = callerMemberName;
            CallerFilePath = callerFilePath;
            CallerLineNumber = callerLineNumber;
            Time = (ticks == 0) ? DateTime.Now : new DateTime(ticks);
            ThreadId = threadId;
            ThreadName = threadName;
            ProcessId = procId;
            AppName = procName;
        }

        public override string ToString()
        {
            return fullMessage ??= FormatMessage(new StringBuilder());
        }

        // Only use in a thread-safe setting!
        internal string ToString(StringBuilder b)
        {
            return fullMessage ??= FormatMessage(b);
        }

        public byte[] ToBytes()
        {
            using var stream = new MemoryStream(1024);
            using var writer = new BinaryWriter(stream);
            writer.Write(ProcessId);
            writer.Write(ThreadId);
            writer.Write(Time.Ticks);
            writer.Write(CallerLineNumber);
            writer.Write((int)Type);
            writer.Write(Message);
            writer.Write(CallerMemberName);
            writer.Write(CallerFilePath);
            writer.Write(AppName);
            writer.Write(ThreadName);
            return stream.ToArray();
        }

        public static LogEntry FromBytes(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes);
            using var reader = new BinaryReader(stream);
            var processId = reader.ReadInt32();
            var threadId = reader.ReadInt32();
            var ticks = reader.ReadInt64();
            var callerLineNumber = reader.ReadInt32();
            var type = (LogMessageType)reader.ReadInt32();
            var message = reader.ReadString();
            var callerMemberName = reader.ReadString();
            var callerFilePath = reader.ReadString();
            var appName = reader.ReadString();
            var threadName = reader.ReadString();

            var entry = new LogEntry(type, ticks, processId, appName, threadId, threadName, message, callerMemberName, callerFilePath, callerLineNumber);
            return entry;
        }


        private string FormatMessage(StringBuilder builder)
        {
            builder.Clear();
            builder.Append(Time.ToString("MMdd HHmm ss.ffffff "));
            builder.Append(LogMessageTypeName[(int)Type]).Append(' ');
            var indent = builder.Length;
            builder.Append("[").Append(AppName).Append(":").Append(ProcessId).Append("] ");
            var tname = string.IsNullOrEmpty(ThreadName) ? ("Thread" + ThreadId) : ThreadName;
            builder.Append("[").Append(tname).Append(":").Append(ThreadId).Append("] ");
            if (builder.Length > 40)
                builder.AppendLine().Append(' ', indent);
            builder.Append(Message ?? "");
            if ((CallerMemberName != null) || (CallerFilePath != null))
            {
                builder.Append(" (");
                if (CallerMemberName != null)
                    builder.Append(CallerMemberName);
                if (CallerFilePath != null)
                {
                    if (CallerMemberName != null)
                        builder.Append(", ");
                    builder.Append(Path.GetFileName(CallerFilePath)).Append(" : ").Append(CallerLineNumber).Append(")");
                }
            }
            return builder.ToString();
        }

        public static string FormatMessage(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            var builder = new StringBuilder();
            builder.Append(DateTime.Now.ToString("MMdd HHmm ss.ffffff "));
            var indent = builder.Length;
            builder.Append("[").Append(Application.Name).Append(":").Append(Process.GetCurrentProcess().Id).Append("] ");
            var tname = string.IsNullOrEmpty(Thread.CurrentThread.Name) ? ("Thread" + Thread.CurrentThread.ManagedThreadId) : Thread.CurrentThread.Name;
            builder.Append("[").Append(tname).Append(":").Append(Thread.CurrentThread.ManagedThreadId).Append("] ");
            if (builder.Length > 40)
                builder.AppendLine().Append(' ', indent);
            builder.Append(message ?? "");
            if ((callerMemberName != null) || (callerFilePath != null))
            {
                builder.Append(" (");
                if (callerMemberName != null)
                    builder.Append(callerMemberName);
                if (callerFilePath != null)
                {
                    if (callerMemberName != null)
                        builder.Append(", ");
                    builder.Append(Path.GetFileName(callerFilePath)).Append(" : ").Append(callerLineNumber).Append(")");
                }
            }
            return builder.ToString();
        }
    }

    #endregion

    #region LogRecorder

    // An object of this class is resposnsible for making a record of some
    // sort for the log entries it receives.
    public abstract class LogRecorder : IDisposable
    {
        private bool disposed;

        #region Concurrency management

        public readonly bool UsingQueue;

        private readonly ConcurrentQueue<LogEntry> pendingEntries;
        private readonly ManualResetEventSlim stopSignal;
        private readonly ManualResetEventSlim stoppedSignal;
        private readonly ManualResetEventSlim readySignal;

        #endregion

        public LogRecorder(bool useQueue)
        {
            UsingQueue = useQueue;

            if (UsingQueue)
            {
                pendingEntries = new ConcurrentQueue<LogEntry>();

                stopSignal = new ManualResetEventSlim(false);
                stoppedSignal = new ManualResetEventSlim(false);
                readySignal = new ManualResetEventSlim(false);

                var thread = new Thread(() => LogThreadFunction());
                thread.Name = "Log";
                thread.Start();
            }
        }

        internal virtual void RecordLogEntry(LogEntry entry)
        {
            if (UsingQueue)
            {
                pendingEntries.Enqueue(entry);
                readySignal.Set();
            }
            else
            {
                RecordLogEntry(entry.ToString());
            }
        }

        // Called from RecordLogEntry(LogEntry), unless the
        // user overrides it.
        protected virtual void RecordLogEntry(string entry)
        {
            var bytes = Encoding.UTF8.GetBytes(entry);
            RecordLogEntry(bytes, 0, bytes.Length);
        }

        // Called from RecordLogEntry(string), unless the
        // user overrides it.
        protected virtual void RecordLogEntry(byte[] bytes, int offset, int count)
        {
            throw new NotImplementedException(); // you have to override at least one form of RecordLogEntry
        }

        #region LogThreadFunction

        protected virtual void LogThreadFunction()
        {
            while (!stopSignal.IsSet)
            {
                readySignal.Wait(1000);
                if (pendingEntries.Count > 0)
                {
                    while (pendingEntries.TryDequeue(out var entry))
                    {
                        try
                        {
                            string s = entry.ToString();
                            RecordLogEntry(s);
                        }
                        catch (ThreadAbortException)
                        {
                            try { stoppedSignal.Set(); } catch { }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                        }
                    }

                    readySignal.Reset(); // probably unnecesary
                }
            }
        }

        #endregion

        #region Dispose

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
                if (disposing)
                {
                }
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }

    #endregion

    #region FileLogRecorder

    public class FileLogRecorder : LogRecorder
    {
        public readonly string LogFolder;
        public readonly string LogFileName;
        protected readonly DateTime StartTime;
        bool disposed;
        protected static readonly byte[] NewLineBytes;
        protected FileStream LogFile;
        public readonly bool AllowsFileSharing;

        static FileLogRecorder()
        {
            NewLineBytes = Encoding.UTF8.GetBytes(Environment.NewLine);
        }

        public FileLogRecorder() : this(true, true, null, null)
        {
        }
        public FileLogRecorder(string logFolder, string appName) : this(true, true, logFolder, appName)
        {
        }

        public FileLogRecorder(bool useQueue, bool allowFileSharing, string logFolder, string appName) : base(useQueue)
        {
            AllowsFileSharing = allowFileSharing;
            StartTime = DateTime.Now;

            logFolder ??= Path.GetTempPath();
            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);

            if (string.IsNullOrWhiteSpace(appName))
                appName = Application.ExeName;

            LogFileName = DetermineFileName(logFolder, appName, StartTime);
            LogFolder = Path.GetDirectoryName(LogFileName);
            Debug.WriteLine(LogEntry.FormatMessage("Using " + LogFileName + " as log"));

            OpenLogFile(LogFileName);
        }

        protected override void RecordLogEntry(byte[] bytes, int offset, int count)
        {
            if (LogFile != null)
            {
                try
                {
                    if (AllowsFileSharing)
                        LogFile.Lock(0, 1);

                    LogFile.Write(bytes, offset, count);
                    LogFile.Write(NewLineBytes, 0, NewLineBytes.Length);
                    LogFile.Flush();

                    if (AllowsFileSharing)
                        LogFile.Unlock(0, 1);
                }
                catch
                {
                    LogFile = null;
                }
            }
        }

        protected virtual void OpenLogFile(string fileName)
        {
            if (AllowsFileSharing)
            {
                LogFile = File.Open(LogFileName, FileMode.OpenOrCreate, FileAccess.Write, FileShare.ReadWrite);
                LogFile.Seek(0, SeekOrigin.End);
                if (LogFile.Position > 0)
                    LogFile.Write(NewLineBytes, 0, NewLineBytes.Length);
            }
            else
            {
                DeleteFile(LogFileName);
                LogFile = File.Open(LogFileName, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
            }
        }

        protected virtual string DetermineFileName(string folder, string appName, DateTime startTime)
        {
            var lfn = Path.Combine(folder, appName + ".log");
            if (!File.Exists(lfn))
                return lfn;

            Debug.WriteLine(LogEntry.FormatMessage("File already exists: " + lfn));
            var end = File.GetLastWriteTime(lfn);
            string backupFilename;

            if (AllowsFileSharing)
            {
                if ((end.DayOfYear == startTime.DayOfYear) && (end.Year == startTime.Year))
                    return lfn;

                var dates = end.ToString("yyyy MMdd");
                backupFilename = Path.Combine(folder, Path.GetFileNameWithoutExtension(lfn) + dates + ".log");
            }
            else
            {
                // Desired file already exists.  Try to back it up, and then reuse it.
                var start = File.GetCreationTime(lfn);

                string dates = start.ToString(" yyyy MMdd HHmmss - ");
                if (start.Year != end.Year)
                    dates += end.ToString("yyyy MMdd HHmmss");
                else if (start.DayOfYear != end.DayOfYear)
                    dates += end.ToString("MMdd HHmmss");
                else
                    dates += end.ToString("HHmmss");

                backupFilename = Path.Combine(folder, Path.GetFileNameWithoutExtension(lfn) + dates + ".log");
            }

            if (File.Exists(backupFilename))
            {
                AppendFileToFile(lfn, backupFilename);
            }
            else
            {
                try
                {
                    File.Move(lfn, backupFilename);
                    Debug.WriteLine(LogEntry.FormatMessage("Existing log file backed up to " + backupFilename));
                }
                catch
                {
                    Debug.WriteLine(LogEntry.FormatMessage("Could not back up  " + lfn));
                }
            }

            return lfn;
        }

        protected static bool DeleteFile(string fn)
        {
            if (!File.Exists(fn))
                return true;
            try
            {
                File.Delete(fn);
                return true;
            }
            catch
            {
                return false;
            }
        }

        protected static bool AppendFileToFile(string sourceFilename, string targetFilename)
        {
            try
            {
                using var b = File.Open(targetFilename, FileMode.Append, FileAccess.ReadWrite);
                using var bsr = new StreamWriter(b);
                b.Seek(0, SeekOrigin.End);
                using var f = File.OpenText(sourceFilename);
                while (true)
                {
                    var line = f.ReadLine();
                    if (line == null)
                        break;
                    bsr.WriteLine(line);
                }
                return true;
            }
            catch
            {
                Debug.WriteLine(LogEntry.FormatMessage("Could not back up  " + sourceFilename));
            }
            return false;
        }

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    LogFile?.Flush();
                    LogFile?.Dispose();
                    LogFile = null;
                    disposed = true;
                }

                base.Dispose(disposing);
            }
        }
    }

    #endregion

    public class NullLogRecorder : LogRecorder
    {
        public NullLogRecorder() : base(false)
        {
        }

        internal override void RecordLogEntry(LogEntry e)
        {
        }
    }

    #region Log

    // The main thing that the Logger class exposes is the MakeLogEntry
    // method.  The other classes in this file use that method to record
    // log entries.
    // Logger has no information about logging levels or other settings;
    // its sole job is to record strings when it is given them.

    public class Log : IDisposable
    {
        public readonly LogRecorder Recorder;
        private bool disposed;
        public readonly LogSettingsCollection Settings;
        public virtual bool Tracking { get; set; }
        private readonly int ProcessId;
        private readonly string AppName;
        [ThreadStatic] protected string ThreadName; // allows us to override Thread.CurrentThread.ThreadName for logging

        public Log(LogRecorder recorder)
        {
            Recorder = recorder;
            Settings = new LogSettingsCollection();
            ProcessId = Process.GetCurrentProcess().Id;
            AppName = Application.Name;
        }

        public Log(Log log)
        {
            Recorder = log.Recorder;
            Settings = new LogSettingsCollection(log.Settings);
            ProcessId = Process.GetCurrentProcess().Id;
            AppName = Application.Name;
        }

        public static implicit operator LogRecorder(Log log) => log.Recorder;

        internal virtual void MakeLogEntry(LogMessageType type, string message, string callerMemberName, string callerFilePath, int callerLineNumber)
        {
            var settings = Settings[type];
            var tid = Thread.CurrentThread.ManagedThreadId;
            ThreadName ??= string.IsNullOrEmpty(Thread.CurrentThread.Name) ? ("Thread" + tid) : Thread.CurrentThread.Name;
            var entry = new LogEntry(type, DateTime.Now.Ticks, ProcessId, AppName, tid, ThreadName, message, callerMemberName, callerFilePath, callerLineNumber);
            if (settings.ToConsole)
                Console.WriteLine(entry.ToString());
            if (settings.ToDebug)
                Debug.WriteLine(entry.ToString());
            if (settings.ToRecorder)
                Recorder.RecordLogEntry(entry);
        }

        #region Dispose

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
                if (disposing)
                {
                    Recorder?.Dispose();
                }
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Trace() methods

        public void Trace(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Debug].NoLogging)
            {
                MakeLogEntry(LogMessageType.Debug, message, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        public void Trace(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Debug].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Debug, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Trace(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!LogSettings[LogMessageType.Debug].NoLogging)
                MakeLogEntry(LogMessageType.Debug, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Info() methods

        public void Info(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Info].NoLogging)
                MakeLogEntry(LogMessageType.Info, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Info(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Info].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Info, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Info(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Info].NoLogging)
                MakeLogEntry(LogMessageType.Info, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Notice() methods

        public void Notice(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Notice].NoLogging)
                MakeLogEntry(LogMessageType.Notice, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Notice(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Notice].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Notice, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Notice(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Notice].NoLogging)
                MakeLogEntry(LogMessageType.Notice, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Warning() methods

        public void Warning(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Warning].NoLogging)
                MakeLogEntry(LogMessageType.Warning, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Warning(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Warning].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Warning, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Warning(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Warning].NoLogging)
                MakeLogEntry(LogMessageType.Warning, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Error() methods

        public void Error(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Error].NoLogging)
                MakeLogEntry(LogMessageType.Error, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Error(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Error].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Error, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Error(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Error].NoLogging)
                MakeLogEntry(LogMessageType.Error, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Critical() methods

        public void Critical(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Critical].NoLogging)
                MakeLogEntry(LogMessageType.Critical, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Critical(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Critical].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Critical, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Critical(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Critical].NoLogging)
                MakeLogEntry(LogMessageType.Critical, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Emergency() methods

        public void Emergency(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Emergency].NoLogging)
                MakeLogEntry(LogMessageType.Emergency, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Emergency(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Emergency].NoLogging)
            {
                var str = ex.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    str += Environment.NewLine + ex.StackTrace;
                MakeLogEntry(LogMessageType.Emergency, str, callerMemberName, callerFilePath, callerLineNumber);
            }
        }

        /*
        public void Emergency(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (!Settings[LogMessageType.Emergency].NoLogging)
                MakeLogEntry(LogMessageType.Emergency, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Write() methods

        public void Write(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            MakeLogEntry(LogMessageType.Info, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Write(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            var str = ex.Message;
            if (!string.IsNullOrEmpty(ex.StackTrace))
                str += Environment.NewLine + ex.StackTrace;
            MakeLogEntry(LogMessageType.Info, str, callerMemberName, callerFilePath, callerLineNumber);
        }

        /*
        public void Log(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            MakeLogEntry(LogMessageType.Info, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion

        #region Tracking methods (Enter(), Exit(), Track())

        public void Enter([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (Tracking && !Settings[LogMessageType.Debug].NoLogging)
                MakeLogEntry(LogMessageType.Debug, "Entering " + callerMemberName, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Exit([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (Tracking && !Settings[LogMessageType.Debug].NoLogging)
                MakeLogEntry(LogMessageType.Debug, "Exiting " + callerMemberName + " at line " + callerLineNumber, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Track([CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (Tracking && !Settings[LogMessageType.Debug].NoLogging)
                MakeLogEntry(LogMessageType.Debug, "At " + callerMemberName + ":" + callerLineNumber, callerMemberName, callerFilePath, callerLineNumber);
        }

        #endregion

        #region Other methods

        public bool Assert(bool expression, /*[CallerArgumentExpression]*/ [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            if (expression)
            {
                if (Tracking && !Settings[LogMessageType.Debug].NoLogging)
                    MakeLogEntry(LogMessageType.Debug, "Assertion passed", callerMemberName, callerFilePath, callerLineNumber);
            }
            else
            {
                if (Tracking && !Settings[LogMessageType.Emergency].NoLogging)
                    MakeLogEntry(LogMessageType.Emergency, "Assertion failed", callerMemberName, callerFilePath, callerLineNumber);
            }
            return expression;
        }

        #endregion
    }

    #endregion Log

    public class LogUser : Log
    {
        public LogUser(LogRecorder recorder) : base(recorder) { }
        public LogUser(Log log) : base(log) { }

        #region Log() methods

        public void Log(string message, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            MakeLogEntry(LogMessageType.Info, message, callerMemberName, callerFilePath, callerLineNumber);
        }

        public void Log(Exception ex, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            var str = ex.Message;
            if (!string.IsNullOrEmpty(ex.StackTrace))
                str += Environment.NewLine + ex.StackTrace;
            MakeLogEntry(LogMessageType.Info, str, callerMemberName, callerFilePath, callerLineNumber);
        }

        /*
        public void Log(object obj, [CallerMemberName] string callerMemberName = null, [CallerFilePath] string callerFilePath = null, [CallerLineNumber] int callerLineNumber = 0)
        {
            MakeLogEntry(LogMessageType.Info, obj?.ToString(), callerMemberName, callerFilePath, callerLineNumber);
        }
        */

        #endregion Log()
    }

    public class NullLog : Log
    {
        public NullLog() : base(new NullLogRecorder())
        {
        }
    }

    public class FileLog : Log
    {
        public string LogFolder => ((FileLogRecorder) Recorder).LogFolder;
        public string LogFileName => ((FileLogRecorder) Recorder).LogFileName;

        public FileLog(string folder, string appName)
            : base(new FileLogRecorder(true, true, folder, appName))
        {
        }

        public FileLog(bool useQueue, bool allowFileSharing, string folder, string appName) 
            : base(new FileLogRecorder(useQueue, allowFileSharing, folder, appName))
        {
        }
    }
}

