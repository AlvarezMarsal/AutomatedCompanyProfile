using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Acp;

namespace Tests
{
    class Program
    {
        static void Main(string[] args)
        {
            TestLocalNode();
        }

        static void TestLocalNode()
        {
            var logger = new FileLogger("Tests");
            var node1 = new LocalNode(logger);
            node1.MessageReceived += (s, e) => TestLocalNodeHandler1(e);
            var node2 = new LocalNode(logger);
            node2.MessageReceived += (s, e) => TestLocalNodeHandler2(e);

            node1.Broadcast("SAY", "TEXT=Hello!");

            node2.SendDatagram(node1.TcpPort, "Hello via Datagram");

            node2.Send(node1.NodeId, "TCP Send/Response Test -- sending");
            Thread.Sleep(10000);
        }

        private static void TestLocalNodeHandler1(LocalNodeMessageEventArgs e)
        {
            Debug.WriteLine("In TestLocalNodeHandler1: received " + e.ToString());
            if (!e.IsUdp)
                e.Respond("TCP Send/Response Test -- responding");
        }

        private static void TestLocalNodeHandler2(LocalNodeMessageEventArgs e)
        {
            Debug.WriteLine("In TestLocalNodeHandler2: received " + e.ToString());
            if (e.IsUdp)
                e.Respond("SAYBACK", "TEXT=Goodbye!");
        }

    }
}
