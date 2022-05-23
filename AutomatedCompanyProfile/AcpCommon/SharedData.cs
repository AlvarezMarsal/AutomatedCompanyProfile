using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public class SharedData : BaseThreaded
    {
        public const long DefaultSize = 65536;
        public const int HeaderSize = sizeof(long) * 2;
        public const string DefaultName = "AcpSharedData";

        public long Size { get; private set; }
        private MemoryMappedFile file;
        private Mutex ownershipMutex;
        private long nextMessagePosition;
        private byte[] readBuffer;
        private byte[] writeBuffer;
        public readonly Action<SharedData, string, string> Handler;
        public string NodeName;

        #region Constructors

        public SharedData(string name, long size, Func<SharedData, string, string, string> func, Logger logger)
            : base(name ?? DefaultName, logger)
        {
            NodeName = Guid.NewGuid().ToString();
            ownershipMutex = new Mutex(false, Name + "Control", out bool _);
            Size = size;
            readBuffer = new byte[4096];
            writeBuffer = new byte[4096];
            Handler = (t, s, m) =>
            {
                var response = func(t, s, m);
                t.Send(s, response);
            };
        }

        public SharedData(string name, Func<SharedData, string, string, string> func, Logger logger)
            : this(name, DefaultSize, func, logger)
        {
        }

        public SharedData(long size, Func<SharedData, string, string, string> func, Logger logger)
            : this(null, size, func, logger)
        {
        }

        public SharedData(Func<SharedData, string, string, string> func, Logger logger)
            : this(null, DefaultSize, func, logger)
        {
        }


        public SharedData(string name, long size, Action<SharedData, string, string> handler, Logger logger)
            : base(name ?? DefaultName, logger)
        {
            NodeName = Guid.NewGuid().ToString();
            Handler = handler;
            ownershipMutex = new Mutex(false, Name + "Control", out bool _);
            Size = size;
            readBuffer = new byte[4096];
            writeBuffer = new byte[4096];
        }

        public SharedData(string name, Action<SharedData, string, string> handler, Logger logger)
            : this(name, DefaultSize, handler, logger)
        {
        }

        public SharedData(long size, Action<SharedData, string, string> handler, Logger logger)
            : this(null, size, handler, logger)
        {
        }

        public SharedData(Action<SharedData, string, string> handler, Logger logger)
            : this(null, DefaultSize, handler, logger)
        {
        }

        #endregion

        // The MMF has the following structure:
        //      bytes 0- 8       the size of the MMF
        //      bytes 9-16       the position of the most recently written message
        //      ...

        protected override void Started()
        {
            base.Started();

            // Set up
            if (!ownershipMutex.WaitOne())
            {
                Stop();
                return;
            }

            try
            {
                var created = false;
                try { file = MemoryMappedFile.OpenExisting(Name); } catch { file = null; }
                if (file == null)
                {
                    created = true;
                    file = MemoryMappedFile.CreateNew(Name, Size);
                }

                using (var view = file.CreateViewAccessor())
                {
                    if (created) // the creator is responsible for writing the first message
                    {
                        nextMessagePosition = HeaderSize; // it's right after the header
                        WriteHeader(view, Size, nextMessagePosition);
                        Log("Created MMF " + Name);
                    }
                    else
                    {
                        Log("Using existing MMF " + Name);
                        ReadHeader(view, out var size, out nextMessagePosition);
                        Size = size;
                    }
                }
            }
            finally
            {
                ownershipMutex.ReleaseMutex();
            }

        }

        protected override void Run()
        {
            // Monitor it for messages
            while (ShouldKeepRunning)
            {
                if (ReadPendingMessages() == 0)
                    Yield(250);
            }
        }

        protected override void Stopping()
        {
            file.Dispose();
        }


        private int ReadPendingMessages()
        {
            var requests = new List<string>();

            ownershipMutex.WaitOne();
            try
            {
                using (var view = file.CreateViewAccessor())
                {
                    //Log("Checking at " + nextMessagePosition);
                    ReadHeader(view, out var size, out var nmp);
                    while (nextMessagePosition != nmp)
                    {
                        Log("Got a message");
                        view.Read(nextMessagePosition, out short length);
                        if (length == short.MinValue) // it wrapped
                        {
                            nextMessagePosition = HeaderSize;
                            view.Read(nextMessagePosition, out length);
                        }

                        if (length > readBuffer.Length)
                            readBuffer = new byte[length * 2];

                        view.ReadArray(nextMessagePosition + sizeof(short), readBuffer, 0, length);
                        nextMessagePosition += sizeof(short) + length;

                        var request = Encoding.UTF8.GetString(readBuffer, 0, length);
                        Log("Read " + request);
                        requests.Add(request);
                    }
                }
            }
            finally
            {
                ownershipMutex.ReleaseMutex();
            }

            foreach (var request in requests)
            {
                var i0 = request.IndexOf('|');
                if (i0 > 0)
                {
                    var target = request.Substring(0, i0);
                    Log("target is " + target);
                    if ((target == "*") || (target == NodeName))
                    {
                        var i1 = request.IndexOf('|', i0 + 1);
                        if (i1 > 0)
                        {
                            var sender = request.Substring(i0 + 1, i1 - i0 - 1);
                            Log("sender is " + sender);
                            var msg = request.Substring(i1 + 1);
                            Log("Handling request (" + sender + " -> " + target + "): " + msg);
                            Handler(this, sender, msg);
                        }
                        else
                        {
                            Warning("Malformed request: (" + i1 + ") " + request);
                        }
                    }
                    else
                    {
                        Log("Ignoring request " + request);
                    }
                }
                else
                {
                    Warning("Malformed request: (" + i0 + ") " + request);
                }
            }

            return requests.Count;
        }

        public void Send(string message)
            => Send(null, message);

        public void SendAll(string message)
            => Send(null, message);

        public void Send(string target, string message)
        {
            // target, sender, body
            var msg = (target ?? "*") + "|" + NodeName + "|" + message;
            var byteCount = Encoding.UTF8.GetByteCount(msg);
            if (byteCount > writeBuffer.Length)
                writeBuffer = new byte[byteCount * 2];
            Encoding.UTF8.GetBytes(msg, 0, msg.Length, writeBuffer, 0);

            Log("Writing " + msg);

            ownershipMutex.WaitOne();
            try
            {
                ReadPendingMessages(); // make sure we've processed all the pending messages
                using (var view = file.CreateViewAccessor())
                {
                    WriteMessage(view, writeBuffer, byteCount);
                }
            }
            finally
            {
                ownershipMutex.ReleaseMutex();
            }

            Log("Wrote at " + nextMessagePosition);
        }

        public void ReadHeader(MemoryMappedViewAccessor view, out long size, out long nmp)
        {
            view.Read(0, out size);
            view.Read(sizeof(long), out nmp);
        }

        public void WriteHeader(MemoryMappedViewAccessor view, long size, long nmp)
        {
            view.Write(0, size);
            view.Write(sizeof(long), nmp);
            view.Flush();
        }

        public void WriteHeader(MemoryMappedViewAccessor view, long nmp)
        {
            view.Write(sizeof(long), nmp);
            view.Flush();
        }

        public void WriteMessage(MemoryMappedViewAccessor view, byte[] bytes, int byteCount)
        {
            // Assumption!  MostRecentMessagePosition is up-to-date!
            // reread the header
            ReadHeader(view, out var size, out var nmp);
            Debug.Assert(nmp == nextMessagePosition);

            // Does it fit?
            var fullLength = sizeof(short) + byteCount;
            var lengthNeeded = fullLength + sizeof(short); // we have to be able to fit at least 1 short AFTER it (for wrapping indicator)
            if ((nextMessagePosition + fullLength) > size)
            {
                // We have to wrap
                view.Write(nextMessagePosition, short.MinValue); // to indicate wrapping
                nextMessagePosition = HeaderSize;
            }

            view.Write(nextMessagePosition, (short) byteCount); // to indicate wrapping
            view.WriteArray(nextMessagePosition + sizeof(short), bytes, 0, byteCount);

            nextMessagePosition += fullLength;
            WriteHeader(view, nextMessagePosition);
        }
    }
}
