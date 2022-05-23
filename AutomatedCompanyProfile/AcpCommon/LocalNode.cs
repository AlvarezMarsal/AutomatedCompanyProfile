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
    // A Local Node is an IPC mechanism for use by processes/threads running on
    // the same machine.
    // It's basically a simple TCP/IP client/server with a UDP mechanism for
    // exchanging port numbers based on names -- each local node has a name.
    // If multiple nodes have the same name, it's pretty random which node
    // will respond.
    //
    // The Local Node allows the sending of messages as strings.
    //
    // Incoming messages trigger LocalNodeMessage event, which the caller
    // can attach to and handle.

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

    public class LocalNode : LogUser, IDisposable
    {
        public event EventHandler<LocalNodeMessageEventArgs> MessageReceived;
        public event EventHandler<LocalNodeExceptionEventArgs> ExceptionOccurred;
        public const int DefaultUdpPort = 17291;
        protected static readonly byte[] EmptyByteArray = new byte[0];
        private const int DefaultBufferSize = 8192;
        public const int DefaultTimeout = int.MaxValue;

        #region ConversationState

        private enum Phase
        {
            AwaitingConnect,
            InitialSend,        // from Send, SendDatagram, or Broadcast
            InitialRead,        // unsolicited message received
            AwaitingResponse,
            Responding,         // sending back a response
            Done
        }

        private class ConversationState
        {
            public LocalNode LocalNode;
            public TcpClient TcpClient;
            public NetworkStream Stream;
            public byte[] Bytes;
            public LocalNodeMessageEventArgs EventArgs;
            public Phase Phase;

            public void Close()
            {
                Phase = Phase.Done;
                Stream?.Close();
                Stream?.Dispose();
                TcpClient?.Close();
            }

            public override string ToString()
            {
                return EventArgs?.ToString() ?? Phase.ToString();
            }
        }


        #endregion

        public readonly int UdpPort;
        public readonly int TcpPort;
        public int NodeId => TcpPort;
        public static readonly int BroadcastNodeId;
        private UdpClient udpClient;
        private bool disposed;
        private TcpListener tcpListener;
        private IPAddress localAddress;
        private static int nextMessageNumber = 0;
        private readonly IPEndPoint broadcastEndPoint;
        private readonly List<TcpClient> openTcpClients;
        public int Timeout;
        public int BufferSize;

        public LocalNode(int udpPort, Log log)
            : base(log)
        {
            UdpPort = udpPort;
            Timeout = DefaultTimeout;
            BufferSize = DefaultBufferSize;
            localAddress = IPAddress.Parse("127.0.0.1");
            broadcastEndPoint = new(IPAddress.Broadcast, UdpPort);

            udpClient = new UdpClient();
            udpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            var ip = new IPEndPoint(IPAddress.Any, UdpPort);
            udpClient.Client.Bind(ip);
            udpClient.Client.SetupBeforeOpen(BufferSize, Timeout);
            var thread = new Thread(() => HandleIncomingUdpMessages());
            thread.Start();
            udpClient.BeginReceive(StaticUdpReceiveCallback, this);
            udpClient.Client.SetupAfterOpen(BufferSize, Timeout);

            openTcpClients = new List<TcpClient>();

            tcpListener = new TcpListener(localAddress, 0);
            tcpListener.Start();
            TcpPort = ((IPEndPoint) tcpListener.LocalEndpoint).Port;
            tcpListener.BeginAcceptTcpClient(StaticTcpAcceptCallback, new ConversationState { LocalNode = this });
        }

        public LocalNode(Log log)
            : this(DefaultUdpPort, log)
        {
        }

        #region Format

        internal byte[] Format(int target, byte[] bytes)
            => Format(target, bytes, 0, bytes.Length);

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

        #region UDP

        private ConcurrentQueue<LocalNodeMessageEventArgs> udpMessageQueue = new ConcurrentQueue<LocalNodeMessageEventArgs>();

        #region Outgoing

        public void Broadcast(byte[] bytes)
            => InternalSendDatagram(Format(BroadcastNodeId, bytes));

        public void Broadcast(byte[] bytes, int index, int count)
            => InternalSendDatagram(Format(BroadcastNodeId, bytes, index, count));

        public void Broadcast(string first, string second = null, string third = null, string fourth = null)
            => InternalSendDatagram(Format(BroadcastNodeId, first, second, third, fourth));

        public void SendDatagram(int target, byte[] bytes)
            => InternalSendDatagram(Format(target, bytes));

        public void SendDatagram(int target, byte[] bytes, int index, int count)
            => InternalSendDatagram(Format(target, bytes, index, count));

        public void SendDatagram(int target, string first, string second = null, string third = null, string fourth = null)
            => InternalSendDatagram(Format(target, first, second, third, fourth));

        private void InternalSendDatagram(byte[] formattedMessage)
        {
            var state = new ConversationState { LocalNode = this, Phase = Phase.InitialSend, Bytes = formattedMessage };
            udpClient?.BeginSend(formattedMessage, formattedMessage.Length, broadcastEndPoint, UdpSendCallback, state);
        }

        private static void UdpSendCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            state.LocalNode?.udpClient?.EndSend(ar);
        }

        #endregion

        #region Incoming

        private static void StaticUdpReceiveCallback(IAsyncResult ar)
        {
            var localNode = (LocalNode) ar.AsyncState;
            localNode.UdpReceiveCallback(ar);
            localNode.udpClient?.BeginReceive(StaticUdpReceiveCallback, localNode);
        }

        private void UdpReceiveCallback(IAsyncResult ar)
        {
            try
            {
                var from = new IPEndPoint(IPAddress.Broadcast, UdpPort);
                var bytes = udpClient?.EndReceive(ar, ref from);
                if (bytes == null)
                {
                    Log("UDP client closed");
                    return;
                }

                var e = Parse(true, bytes);
                if (IsForThisNode(e))
                {
                    try
                    {
                        udpMessageQueue.Enqueue(e);
                    }
                    catch (Exception ex)
                    {
                        Log(ex);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        private void HandleIncomingUdpMessages()
        {
            while (true)
            {
                try
                {
                    while (udpMessageQueue.TryDequeue(out var e))
                    {
                        MessageReceived?.Invoke(this, e);
                        if (e.FormattedResponse != null)
                            udpClient.Send(e.FormattedResponse, e.FormattedResponse.Length, broadcastEndPoint);
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

        #endregion

        #region TCP/IP

        #region Outgoing

        public void Send(int target, byte[] bytes)
            => InternalSend(false, target, Format(target, bytes));
        public void Send(bool expectResponse, int target, byte[] bytes)
             => InternalSend(expectResponse, target, Format(target, bytes));

        public void Send(int target, byte[] bytes, int index, int count)
            => InternalSend(false, target, Format(target, bytes, index, count));
        public void Send(bool expectResponse, int target, byte[] bytes, int index, int count)
            => InternalSend(expectResponse, target, Format(target, bytes, index, count));

        public void Send(int target, string first, string second = null, string third = null, string fourth = null)
            => InternalSend(false, target, Format(target, first, second, third, fourth));
        public void Send(bool expectResponse, int target, string first, string second = null, string third = null, string fourth = null)
            => InternalSend(expectResponse, target, Format(target, first, second, third, fourth));

        private void InternalSend(bool expectResponse, int target, byte[] formattedMessage)
        {
            var tcpClient = new TcpClient();
            tcpClient.ReceiveTimeout = Timeout;
            tcpClient.SendTimeout = Timeout;
            var state = new ConversationState { LocalNode = this, TcpClient = tcpClient, Bytes = formattedMessage, Phase = Phase.AwaitingConnect };
            tcpClient.BeginConnect("127.0.0.1", target, TcpConnectCallback, state);
        }

        // This function is called when we establish a TCP/IP connection to thw
        // target.
        // It starts the process of sending the message.
        private static void TcpConnectCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            if (state.TcpClient != null)
            {
                state.TcpClient.EndConnect(ar);
                state.Stream = state.TcpClient.GetStream();
                if (state.Stream != null)
                {
                    state.Phase = Phase.InitialSend;
                    state.Stream.BeginWrite(state.Bytes, 0, state.Bytes.Length, TcpWriteCallback, state);
                }
            }
        }

        // This function is called when we have finished sending the message.
        // If this is a fire-and-forget situattion, we close the connection and are done.
        // Otherwise, we start the process of reading the response.
        private static void TcpWriteCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            if (state?.Stream != null)
            {
                state.Stream.EndWrite(ar);
                if (state.Phase == Phase.Responding)
                {
                    state.Close(); // we always do a simple open -> request -> response -> close
                }
                else // prepare to get a response
                {
                    if (state.Bytes.Length < state.LocalNode.BufferSize)
                        state.Bytes = new byte[state.LocalNode.BufferSize];
                    state.Phase = Phase.AwaitingResponse;
                    state.Stream.BeginRead(state.Bytes, 0, state.Bytes.Length, TcpReadCallback, state);
                }
            }
        }

        #endregion

        #region Incoming

        private static void StaticTcpAcceptCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            state.LocalNode.TcpAcceptCallback(ar);
        }

        private void TcpAcceptCallback(IAsyncResult ar)
        {
            if (tcpListener == null)
                return;

            try
            {
                using var tcpClient = tcpListener.EndAcceptTcpClient(ar);
                tcpClient.Client.Setup(BufferSize, Timeout);

                tcpListener.BeginAcceptTcpClient(StaticTcpAcceptCallback, new ConversationState { LocalNode = this });

                var bytes = new byte[tcpClient.Client.ReceiveBufferSize];
                using var stream = tcpClient.GetStream();

                while (true)
                {
                    int length = stream.Read(bytes, 0, bytes.Length);
                    if (length < 1)
                    {
                        Log("Client closed");
                        break;
                    }

                    if (length == bytes.Length)
                        Log("Message MIGHT be too long");

                    var handler = MessageReceived;
                    if (handler == null)
                    {
                        Log("No handler for message");
                    }
                    else
                    {
                        var e = Parse(false, bytes, 0, length);
                        if (e != null)
                        {
                            handler.Invoke(this, e);

                            if (e.FormattedResponse == null)
                            {
                                Log("No response");
                                break;
                            }

                            Log("Sending response");
                            stream.Write(e.FormattedResponse, 0, e.FormattedResponse.Length);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        // This method is called when we have received a response.
        // We pass the parsed message to whataver event handler is
        // hooked up.  That handler can issue a response of its own;
        // if so, we send it.
        // Otherwise, we close the socket down.
        private static void TcpReadCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            if (state?.Stream != null)
            {
                int length = state.Stream.EndRead(ar);
                if (length < 1)
                {
                    state.Close();
                }
                else
                {
                    state.EventArgs = state.LocalNode.Parse(false, state.Bytes, 0, length);
                    if (state.LocalNode.IsForThisNode(state.EventArgs))
                    {
                        state.LocalNode.MessageReceived?.BeginInvoke(state.LocalNode, state.EventArgs, TcpMessageReceivedCallback, state);
                    }
                }
            }
        }

        // We call this after we have fully received the response, to allow
        // the user code to act on the response.
        private static void TcpMessageReceivedCallback(IAsyncResult ar)
        {
            var state = (ConversationState) ar.AsyncState;
            var h = state?.LocalNode?.MessageReceived;
            if (h != null)
            {
                h.EndInvoke(ar);

                if (state.EventArgs.FormattedResponse == null)
                {
                    // No response was made
                    state.Close();
                }
                else
                {
                    // If the handler issued a further response, we send that to the target
                    state.Phase = Phase.Responding;
                    state.Stream.BeginWrite(state.EventArgs.FormattedResponse, 0, state.EventArgs.FormattedResponse.Length, TcpWriteCallback, state);
                }
            }
        }

        #endregion

        #endregion

        #region Fire an exception

        private void ThrowException(Exception ex)
        {
            var e = new LocalNodeExceptionEventArgs(this, ex);
            ExceptionOccurred?.Invoke(this, e);
            if (!e.Handled)
                throw e.Exception;  
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
