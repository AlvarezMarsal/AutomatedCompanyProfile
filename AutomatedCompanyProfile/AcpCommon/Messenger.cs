using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public interface IMessageParser<M>
    {
        M Parse(byte[] raw, int offset, int count);
    }

    public interface IMessageFormatter<M>
    {
        byte[] Format(M msg);
        int Format(M msg, ref byte[] bytes);
    }

    public abstract class MessageFormatter<M> : IMessageParser<M>, IMessageFormatter<M>
    {
        private int nextMessageNumber = -1;
        protected int NextMessageNumber => Interlocked.Increment(ref nextMessageNumber);

        public byte[] Format(M msg)
        {
            byte[] bytes = null;
            Format(msg, ref bytes);
            return bytes;
        }

        public abstract int Format(M msg, ref byte[] bytes);

        public abstract M Parse(byte[] raw, int offset, int count);
    }

    public class StringMessageFormatter : MessageFormatter<string>
    {
        // [ThreadStatic] byte[] buffer;

        public override int Format(string msg, ref byte[] bytes)
        {
            int length = Encoding.UTF8.GetByteCount(msg);
            if ((bytes == null) || (bytes.Length < length))
            {
                bytes = new byte[length];
            }

            Encoding.UTF8.GetBytes(msg, 0, msg.Length, bytes, 0);
            return length;
        }

        public override string Parse(byte[] raw, int offset, int count)
        {
            return Encoding.UTF8.GetString(raw, offset, count);
        }

    }

    public class MessageReceivedEventArgs<M> : EventArgs
    {
        public readonly M Message;
    }


    // Handles incoming/outgoing messages
    public class Messenger : BaseThread
    {
        private readonly StringMessageFormatter formatter = new StringMessageFormatter();
        public readonly int Timeout;
        public readonly int BufferSize;
        public const int DefaultTimeout = int.MaxValue;
        public const int DefaultBufferSize = 8192;
        public int Port { get; private set; }
        public Func<string, string> Handler { get; private set; }
        // ConcurrentDictionary<int, Client> clients = new ConcurrentDictionary<int, Client>();
        // int nextClientId = 0;
        TcpListener tcpListener;

        public Messenger(int port, Func<string, string> handler, Log log)
            : base("Messenger"+port, log)
        {
            Port = port;
            Handler = handler;
            Timeout = DefaultTimeout;
            BufferSize = DefaultBufferSize;
        }

        public Messenger(Func<string, string> handler, Log log)
            : this(0, handler, log)
        {
        }

        protected override void Started()
        {
            base.Started();

            var localAddr = IPAddress.Parse("127.0.0.1");
            tcpListener = new TcpListener(localAddr, Port);
            tcpListener.Start();

            if (Port == 0)
            {
                Port = ((IPEndPoint) tcpListener.LocalEndpoint).Port;
                Name = "Messenger" + Port;
            }
        }

        protected override void Run()
        {
            Info("Waiting for TCP messages on port " + Port);

            while (!StopSignalReceived)
            {
                if (tcpListener.Pending())
                {
                    Trace("Detected incoming message");
                    var tcpClient = tcpListener.AcceptTcpClient();
                    tcpClient.Client.Setup(BufferSize, Timeout);
                    var thread = new Thread(() => ServiceClient(tcpClient));
                    thread.Start();
                }
                else
                {
                    Yield(100);
                }
            }

            tcpListener.Stop();
            tcpListener = null;

            Info("Messenger stopping");
        }

        private void ServiceClient(TcpClient tcpClient)
        {
            try
            {
                var buffer = new byte[tcpClient.ReceiveBufferSize];
                using var stream = tcpClient.GetStream();

                var length = stream.Read(buffer, 0, buffer.Length);
                if (length < 1)
                {
                    Log("Stream closed");
                    return;
                }

                var h = Handler;
                if (h == null)
                {
                    Log("No handler for request");
                }
                else
                {
                    var request = formatter.Parse(buffer, 0, length);
                    var response = h(request);
                    if (response == null)
                    {
                        Log("No response to send");
                        return;
                    }
                    else
                    {
                        length = formatter.Format(response, ref buffer);
                        stream.Write(buffer, 0, length);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new AcpException(ex);
            }
            finally
            {
                tcpClient.Dispose();
            }
        }

        public string Send(int targetPort, string msg, bool expectResponse = true)
        {
            try
            {
                using var tcpClient = new TcpClient("127.0.0.1", targetPort);
                tcpClient.Client.Setup(BufferSize, Timeout);
                tcpClient.Client.DontFragment = true;

                var buffer = new byte[tcpClient.Client.ReceiveBufferSize];
                var length = formatter.Format(msg, ref buffer);
                using var stream = tcpClient.GetStream();
                stream.Write(buffer, 0, length);

                if (!expectResponse)
                    return null;

                length = stream.Read(buffer, 0, buffer.Length);
                if (length < 1)
                {
                    Log("Socket closed");
                    return "";
                }

                if (length == buffer.Length)
                    Warning("Response MIGHT be too long");

                var response = formatter.Parse(buffer, 0, length);
                return response;
            }
            catch (Exception ex)
            {
                Log(ex);
                return "ERROR=" + ex.Message;
            }
        }
    }
}
