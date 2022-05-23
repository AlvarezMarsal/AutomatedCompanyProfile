using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Configuration;
using System.Reflection;

namespace Acp
{
    public class Server : Threaded, IDisposable
    {
        private static readonly Counter Counter = new (DateTime.Now.Ticks);
        public bool TryToCleanCache { get; set; } = true;
        public string OutputFolder { get; private set; }
        public string InputFolder { get; private set; }
        private const string CiqInterfaceFilename = "CIQInterface.xlsm";

        public Server(Log log) : base("Server", log)
        {
        }

        protected override void Started()
        {
            if (TryToCleanCache)
                CleanCache();
        }

        #region Clean up the cache

        private bool CleanCache()
        {
            Info("Cleaning cache");
            try
            {
                var cacheFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                    "Capital IQ",
                                    "Office Plug-in");
                if (Directory.Exists(cacheFolder))
                {
                    // Only the first ACP Server will be able to delete the db
                    File.Delete(Path.Combine(cacheFolder, "ciqudb.sdf"));
                    var md = Path.Combine(cacheFolder, "MetaData");
                    if (Directory.Exists(md))
                        Directory.Delete(md, true);
                    Info("Cleaned cache");
                }

                return true;
            }
            catch (Exception)
            {
                Warning("Could not clean cache");
                return false;
            }
        }

        #endregion

        protected override void Run()
        {
            Client client = null;

            string rootOutputFolder = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Output");
            if (!Directory.Exists(rootOutputFolder))
                Directory.CreateDirectory(rootOutputFolder);

            OutputFolder = Path.Combine(rootOutputFolder, DateTime.Now.ToString("yyyy-MM-dd"));
            if (!Directory.Exists(OutputFolder))
                Directory.CreateDirectory(OutputFolder);

            InputFolder = ConfigurationManager.AppSettings.Get(Environment.MachineName + "-Input");
            if (!Directory.Exists(InputFolder))
                Directory.CreateDirectory(InputFolder);

            while (!StopSignalReceived)
            {
                if (Application.Signals["CiqError"].IsSet)
                    break;

                if (client == null)
                {
                    Log("Constructing client");

                    var cifn = MapInputFilename(CiqInterfaceFilename);
                    var workbookFilename = MapOutputFilename(Application.ExeName + "-" + Counter.Next + "-ciq.xlsm");
                    File.Copy(cifn, workbookFilename, true);
                    var a = File.GetAttributes(workbookFilename);
                    File.SetAttributes(workbookFilename, a & ~FileAttributes.ReadOnly);

                    client = new Client(this, workbookFilename);
                    client.Start();
                }
                else
                {
                    Yield(100);
                }
            }

            client?.Stop();
        }

        public string MapInputFilename(string filename)
        {
            var fullFilename = Path.Combine(InputFolder, filename);
            Debug.Assert(File.Exists(fullFilename));
            return fullFilename;
        }

        public string MapOutputFilename(string filename)
        {
            var fullFilename = Path.Combine(OutputFolder, filename);
            return fullFilename;
        }

    }
}
