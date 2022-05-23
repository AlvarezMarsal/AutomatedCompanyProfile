using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    /*
    #region LocalNodeMessageEventArgs

    public class LocalNodeMessageEventArgs : EventArgs
    {
        public readonly LocalNode LocalNode;
        public bool IsBroadcast => (Target == 0);
        public readonly bool IsUdp;
        public readonly int Target;
        public readonly int Sender;
        public readonly byte[] MessageBytes;
        public readonly int MessageNumber;
        private string[] message;
        public bool KeepOpen;
        public string[] Message => message ??= ParseMessageBytes();
        internal byte[] FormattedResponse;

        internal LocalNodeMessageEventArgs(LocalNode node, bool isUdp, int target, int sender, int messageNumber, byte[] bytes)
        {
            IsUdp = isUdp;
            LocalNode = node;
            Target = target;
            Sender = sender;
            MessageNumber = messageNumber;
            MessageBytes = bytes;
        }

        public void Respond(byte[] bytes)
            => FormattedResponse = LocalNode.Format(Sender, bytes);

        public void Respond(byte[] bytes, int index, int count)
            => FormattedResponse = LocalNode.Format(Sender, bytes, index, count);

        public void Respond(string first, string second = null, string third = null, string fourth = null)
            => FormattedResponse = LocalNode.Format(Sender, first, second, third, fourth);

        private string[] ParseMessageBytes()
        {
            var list = new List<string>();

            using var stream = new MemoryStream(MessageBytes);
            using var reader = new BinaryReader(stream);
            while (reader.PeekChar() != -1)
            {
                list.Add(reader.ReadString());
            }
            return list.ToArray();
        }

        public override string ToString()
        {
            var b = new StringBuilder();
            b.Append('[').Append(Sender).Append("->").Append(Target).Append("] ");
            b.Append(MessageBytes.Length).Append(' ').Append("bytes ");
            foreach (var m in Message)
            {
                b.Append('\"').Append(m).Append("\" ");
            }

            if (FormattedResponse != null)
            {
                var e = LocalNode.Parse(IsUdp, FormattedResponse);
                if (e != null)
                    b.AppendLine().Append("    " + e.ToString());
            }
            return b.ToString();
        }

    }

    #endregion

    #region LocalNodeExceptionEventArgs

    public class LocalNodeExceptionEventArgs : EventArgs
    {
        public readonly LocalNode LocalNode;
        public readonly Exception Exception;
        public bool Handled = false;

        internal LocalNodeExceptionEventArgs(LocalNode node, Exception ex)
        {
            LocalNode = node;
            Exception = ex;
        }

        internal LocalNodeExceptionEventArgs(LocalNode node, string msg, Exception ex)
        {
            LocalNode = node;
            Exception = new Exception(msg, ex);
        }

        internal LocalNodeExceptionEventArgs(LocalNode node, string msg)
        {
            LocalNode = node;
            Exception = new Exception(msg);
        }

        public override string ToString()
            => Exception.ToString();
    }

    #endregion
    */

    public class UdpNodeMessage
    {
        public int TotalLength;
        public int Target;
        public int Sender;
        public int MessageNumber;
        public readonly List<byte[]> ByteSequences = new List<byte[]>();
    }

    public class UdpMessageReceivedEventArgs : MessageReceivedEventArgs<UdpNodeMessage>
    {
    }

    public class UdpNodeMessageFormatter : MessageFormatter<UdpNodeMessage>
    {
        private readonly int Sender;

        public UdpNodeMessageFormatter(int senderPort)
        {
            Sender = senderPort;
        }

        public override int Format(UdpNodeMessage msg, ref byte[] bytes)
        {
            var totalLength = sizeof(int) + sizeof(int) + sizeof(int) + sizeof(int);
 
            foreach (var seq in msg.ByteSequences)
                totalLength += sizeof(int) + seq.Length;

            if ((bytes == null) || (bytes.Length < totalLength))
                bytes = new byte[totalLength];

            using var stream = new MemoryStream(bytes, 0, totalLength, true);
            using var writer = new BinaryWriter(stream);
            writer.Write(totalLength);
            writer.Write(msg.Target);
            writer.Write(Sender);
            writer.Write(NextMessageNumber);

            foreach (var seq in msg.ByteSequences)
                writer.Write(seq);

            Debug.Assert(stream.Length == totalLength);
            totalLength = (int) stream.Length;
            return totalLength;
        }

        public override UdpNodeMessage Parse(byte[] bytes, int offset, int count)
        {
            var message = new UdpNodeMessage();
            using var stream = new MemoryStream(bytes, offset, count, false);
            using var reader = new BinaryReader(stream);
            message.TotalLength = reader.ReadInt32();
            Debug.Assert(stream.Length == message.TotalLength);
            message.Target = reader.ReadInt32();
            message.Sender = reader.ReadInt32();

            while (reader.PeekChar() != -1)
            {
                var len = reader.ReadInt32();
                var seq = reader.ReadBytes(len);
                message.ByteSequences.Add(seq);
            }

            return message;
        }
    }

    public class UdpNode : LogUser, IDisposable
    {
        public const int DefaultBroadcastPort = 17291;
        protected readonly int BroadcastPort;
        public int BroadcastNodeId => BroadcastPort;

        private const int DefaultBufferSize = 8192;
        protected readonly int BufferSize;

        public const int DefaultTimeout = int.MaxValue;
        public readonly int Timeout;

        public event EventHandler<UdpMessageReceivedEventArgs> MessageReceived;
        //     public event EventHandler<LocalNodeExceptionEventArgs> ExceptionOccurred;
        protected static readonly byte[] EmptyByteArray = new byte[0];
        public readonly int Port;
        private IPAddress localAddress;
        private readonly IPEndPoint broadcastEndPoint;
        private UdpClient broadcastClient;
        private UdpClient client;
        private readonly UdpNodeMessageFormatter formatter;
        public int NodeId => Port;

        /*
        public readonly int TcpPort;
        public int NodeId => TcpPort;
        public static readonly int BroadcastNodeId;
        private bool disposed;
        private TcpListener tcpListener;
        private static int nextMessageNumber = 0;
        private readonly List<TcpClient> openTcpClients;
        */
        private readonly ConcurrentQueue<byte[]> messageQueue;

        #region Constructors

        public UdpNode(int broadcastUdpPort, int bufferSize, int timeout, Logger logger)
            : base(logger)
        {
            BroadcastPort = broadcastUdpPort;
            Timeout = Math.Max(100, timeout);
            BufferSize = Math.Max(8192, bufferSize);
            localAddress = IPAddress.Parse("127.0.0.1");
            broadcastEndPoint = new(IPAddress.Broadcast, BroadcastPort);

            messageQueue = new ConcurrentQueue<byte[]>();

            broadcastClient = new UdpClient();
            broadcastClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            var ip = new IPEndPoint(IPAddress.Any, BroadcastPort);
            broadcastClient.Client.Bind(ip);
            broadcastClient.Client.Setup(BufferSize, Timeout);

            client = new UdpClient();
            ip = new IPEndPoint(IPAddress.Any, 0);
            client.Client.Bind(ip);
            client.Client.Setup(BufferSize, Timeout);
            Port = ip.Port;

            formatter = new UdpNodeMessageFormatter(Port);

            var thread = new Thread(() => HandleIncomingUdpMessages());
            thread.Start();

            broadcastClient.BeginReceive(StaticBroadcastReceiveCallback, this);
            client.BeginReceive(StaticReceiveCallback, this);
        }

        public UdpNode(int bufferSize, int timeout, Logger logger)
            : this(DefaultBroadcastPort, bufferSize, timeout, logger)
        {
        }

        public UdpNode(int broadcastUdpPort, Logger logger)
            : this(0, DefaultBufferSize, DefaultTimeout, logger)
        {
        }

        public UdpNode(Logger logger)
            : this(DefaultBroadcastPort, DefaultBufferSize, DefaultTimeout, logger)
        {
        }

        #endregion

        #region Format

        internal byte[] Format(byte[] bytes)
            => Format(bytes, 0, bytes.Length);

        internal byte[] Format(string str)
            => Format(bytes, 0, bytes.Length);

        internal byte[] Format(int target, byte[] bytes, int index, int count)
        {
            using var stream = new MemoryStream();
            using var writer = new BinaryWriter(stream);
            int totalLength = 0;
            writer.Write(totalLength);
            writer.Write(target);
            writer.Write(NodeId);
            writer.Write(Interlocked.Increment(ref nextMessageNumber));
            writer.Write(bytes, index, count);
            totalLength = (int)writer.Seek(0, SeekOrigin.Current);
            writer.Seek(0, SeekOrigin.Begin);
            writer.Write(totalLength);
            return stream.ToArray();
        }

        internal byte[] Format(int target, string first, string second = null, string third = null, string fourth = null)
        {
            using var stream = new MemoryStream();
            using var writer = new BinaryWriter(stream);
            int totalLength = 0;
            writer.Write(totalLength);
            writer.Write(target);
            writer.Write(NodeId);
            writer.Write(Interlocked.Increment(ref nextMessageNumber));
            writer.Write(first ?? "");
            if (second != null)
                writer.Write(second);
            if (third != null)
                writer.Write(third);
            if (fourth != null)
                writer.Write(fourth);
            totalLength = (int) writer.Seek(0, SeekOrigin.Current);
            writer.Seek(0, SeekOrigin.Begin);
            writer.Write(totalLength);
            return stream.ToArray();
        }

        internal LocalNodeMessageEventArgs Parse(bool isUdp, byte[] raw)
            => Parse(isUdp, raw ?? EmptyByteArray, 0, raw?.Length ?? 0);

        internal LocalNodeMessageEventArgs Parse(bool isUdp, byte[] raw, int offset, int length)
        {
            if (length < 16) // too short to have a proper header
                return null;

            using var stream = new MemoryStream(raw, offset, length);
            using var reader = new BinaryReader(stream);

            var totalLength = reader.ReadInt32();
            var target = reader.ReadInt32();
            var sender = reader.ReadInt32();
            var messageNumber = reader.ReadInt32();
            var bytes = reader.ReadBytes(totalLength - sizeof(int) - sizeof(int));

            return new LocalNodeMessageEventArgs(this, isUdp, target, sender, nextMessageNumber, bytes);
        }

        #endregion

        #region Outgoing

        // Local broadcast
        public void Broadcast(byte[] bytes)
        {
            var msg = new UdpNodeMessage { Target = BroadcastNodeId };
            msg.ByteSequences.Add(bytes);
            InternalBroadcast(msg);
        }

        public void Broadcast(byte[] bytes, int offset, int count)
        {
            var msg = new UdpNodeMessage { Target = BroadcastNodeId };
            if ((offset == 0) && (count == bytes.Length))
            {
                msg.ByteSequences.Add(bytes);
            }
            else
            {
                var copy = new byte[count];
                Array.Copy(bytes, offset, copy, 0, count);
            }
            InternalBroadcast(msg);
        }

        public void Broadcast(string first, string second = null, string third = null, string fourth = null)
        {
            var msg = new UdpNodeMessage { Target = BroadcastNodeId };
            msg.ByteSequences.Add(Encoding.UTF8.GetBytes(first));
            if (second != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(second));
            if (third != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(third));
            if (fourth != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(fourth));
            InternalBroadcast(msg);
        }

        private void InternalBroadcast(UdpNodeMessage msg)
        {
            var bytes = formatter.Format(msg);
            broadcastClient?.Send(bytes, bytes.Length, broadcastEndPoint);
        }


        public void Send(int target, byte[] bytes)
        {
            var msg = new UdpNodeMessage { Target = target };
            msg.ByteSequences.Add(bytes);
            InternalSend(target, msg);
        }

        public void Send(int target, byte[] bytes, int offset, int count)
        {
            var msg = new UdpNodeMessage { Target = target };
            if ((offset == 0) && (count == bytes.Length))
            {
                msg.ByteSequences.Add(bytes);
            }
            else
            {
                var copy = new byte[count];
                Array.Copy(bytes, offset, copy, 0, count);
            }
            InternalSend(target, msg);
        }

        public void Send(int target, string first, string second = null, string third = null, string fourth = null)
        {
            var msg = new UdpNodeMessage { Target = target };
            msg.ByteSequences.Add(Encoding.UTF8.GetBytes(first));
            if (second != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(second));
            if (third != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(third));
            if (fourth != null)
                msg.ByteSequences.Add(Encoding.UTF8.GetBytes(fourth));
            InternalSend(target, msg);
        }

        private void InternalSend(int target, UdpNodeMessage msg)
        {
            msg.Target = target;
            var bytes = formatter.Format(msg);
            client?.Send(bytes, bytes.Length, new IPEndPoint(localAddress, target));
        }

        #endregion

        #region Incoming

        private static void StaticReceiveCallback(IAsyncResult ar)
        {
            var node = (UdpNode) ar.AsyncState;
            var from = new IPEndPoint(IPAddress.Broadcast, node.Port);
            var bytes = node.client.EndReceive(ar, ref from);
            node.messageQueue.Enqueue(bytes);
            node.client.BeginReceive(StaticReceiveCallback, node);
        }

        private static void StaticBroadcastReceiveCallback(IAsyncResult ar)
        {
            var node = (UdpNode)ar.AsyncState;
            var from = new IPEndPoint(IPAddress.Broadcast, node.BroadcastPort);
            var bytes = node.broadcastClient.EndReceive(ar, ref from);
            node.messageQueue.Enqueue(bytes);
            node.broadcastClient.BeginReceive(StaticReceiveCallback, node);
        }

        private void HandleIncomingUdpMessages()
        {
            while (true)
            {
                try
                {
                    while (messageQueue.TryDequeue(out var bytes))
                    {
                        var msg = formatter.Parse(bytes, 0, bytes.Length);
                        var e = new UdpMessageReceivedEventArgs(msg);
                        MessageReceived?.Invoke(this, e);
                        if (e.FormattedResponse != null)
                            client.Send(e.FormattedResponse, e.FormattedResponse.Length, broadcastEndPoint);
                    }
                }
                catch (Exception ex)
                {
                    Log(ex);
                }
            }
        }


        private bool IsForThisNode(LocalNodeMessageEventArgs e)
        {
            if (e == null)
                return false;
            if (e.Target == NodeId)
                return true;
            return (e.Target == BroadcastNodeId) && (e.Sender != NodeId);
        }

        private static void UdpMessageReceivedCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            try
            {
                state.LocalNode?.MessageReceived?.EndInvoke(ar);

             }
            catch (AcpException ax)
            {
                state.LocalNode?.Warning(ax);
                state.LocalNode?.Dispose();
                throw;
            }
            catch (Exception ex)
            {
                state.LocalNode?.Warning(ex);
                state.LocalNode?.Dispose();
                throw;
            }

        }

        #endregion

        #region Dispose

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    udpClient?.Close();
                    udpClient = null;
                    tcpListener?.Stop();
                    tcpListener = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposed = true;
                base.Dispose(disposing);
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~LocalNode()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }


        #endregion
    }
}
