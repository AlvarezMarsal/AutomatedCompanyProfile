using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Maintenance
{
    class Program
    {
        static object tcpClientsLock = new object();
        static List<TcpClient> tcpClients = new List<TcpClient>();

        static void Main(string[] args)
        {
            var port = 1235;
            var localAddr = IPAddress.Parse("127.0.0.1");
            var tcpListener = new TcpListener(localAddr, port);
            tcpListener.Start();

            Console.WriteLine("Hit ESC to stop the maintenance app");

            while (true)
            {
                if (tcpListener.Pending())
                {
                    var tcpClient = tcpListener.AcceptTcpClient();
                    if (tcpListener.Pending())
                    {
                        Console.WriteLine("Client is connecting");
                        tcpClient = tcpListener.AcceptTcpClient();
                        lock (tcpClientsLock)
                            tcpClients.Add(tcpClient);
                        StartClientRead(tcpClient);
                    }
                }
                else if (Console.KeyAvailable)
                {
                    var k = Console.ReadKey(true);
                    if (k.Key == ConsoleKey.Escape)
                    {
                        Console.WriteLine("User has requested a shutdown");
                        break;
                    }
                }
                else
                {
                    Thread.Sleep(250);
                }
            }

            tcpListener.Stop();
            tcpListener = null;

            lock (tcpClientsLock)
            {
                foreach (var c in tcpClients)
                    StopClient(c, null);
                tcpClients.Clear();
            }
        }

        private static void StartClientRead(TcpClient tcpClient)
        {
            var stream = tcpClient.GetStream();
            stream.ReadTimeout = 30 * 60 * 1000;
            var buffer = new byte[4096];
            try
            {
                var result = stream.BeginRead(buffer, 0, buffer.Length, ClientReadCallback, Tuple.Create(buffer, tcpClient, stream));
            }
            catch
            {
                Console.WriteLine("Client is disconnecting");
                StopClient(tcpClient, stream);
            }
        }

        private static void ClientReadCallback(IAsyncResult ar)
        {
            var tuple = (Tuple<byte[], TcpClient, NetworkStream>) ar.AsyncState;
            var buffer = tuple.Item1;
            var tcpClient = tuple.Item2;
            var stream = tuple.Item3;

            var length = stream.EndRead(ar);
            if (length < 1)
            {
                StopClient(tcpClient, stream);
                return;
            }

            var bytes = Encoding.UTF8.GetBytes("ERROR=The ACP Server is down for maintenance.");
            stream.Write(bytes, 0, bytes.Length);

            try
            {
                var result = stream.BeginRead(buffer, 0, buffer.Length, ClientReadCallback, Tuple.Create(buffer, tcpClient, stream));
            }
            catch
            {
                Console.WriteLine("Client is disconnecting");
                StopClient(tcpClient, stream);
            }
        }

        private static void StopClient(TcpClient tcpClient, NetworkStream stream)
        {
            stream?.Close();
            stream = null;
            lock (tcpClientsLock)
                tcpClients.Remove(tcpClient);
            tcpClient.Close();
            tcpClient = null;
        }
    }
}
