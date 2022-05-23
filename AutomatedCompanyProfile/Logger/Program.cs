using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;

namespace Logger
{
    class Program
    {
        static void Main(string[] args)
        {
            var logger = new Acp.QueuedFileLogger(null, "AcpLog"); // the actual log-to-file guy
            var logger2 = new Acp.NullLogger(); 
            var logger3 = new Acp.NullLogger();
            var queue = new ConcurrentQueue<Acp.LocalNodeMessageEventArgs>();
            var localNode = new Acp.LocalNode(logger2);
            localNode.MessageReceived += (_, e) => HandleIncomingMessage(localNode, queue, e);

            var lt = new DequeueThread(queue, logger, logger3);
            lt.Start();

            while (!lt.IsStopped)
                Thread.Sleep(100);

            logger.Dispose();
        }

        static void HandleIncomingMessage(Acp.LocalNode node, ConcurrentQueue<Acp.LocalNodeMessageEventArgs> queue, Acp.LocalNodeMessageEventArgs e)
        {
            if (e.Message[0] == "FINDLOGGER")
                e.Respond("LOGGER", node.NodeId.ToString());
            else if (e.Message[0] == "LOG")
                queue.Enqueue(e);
        }

        class DequeueThread : Acp.BaseThread
        {
            ConcurrentQueue<Acp.LocalNodeMessageEventArgs> queue;
            Acp.Logger realLogger;

            public DequeueThread(ConcurrentQueue<Acp.LocalNodeMessageEventArgs> queue,
                                 Acp.Logger realLogger,
                                 Acp.Logger fakeLogger) : base("AcpLogDequeue", fakeLogger)
            {
                this.queue = queue;
                this.realLogger = realLogger;
            }

            protected override void Run()
            {
                while (!IsStopped)
                {
                    while (queue.TryDequeue(out var msg))
                    {
                        realLogger.RecordLogEntry(msg.);
                    }
                }
            }
        }
    }
}
