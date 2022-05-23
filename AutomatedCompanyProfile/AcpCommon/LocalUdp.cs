using System;
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

    #region UdpNodeMessageEventArgs

    public class UdpMessageEventArgs : MessageEventArgs<int>
    {
        public override bool IsBroadcast => (Target == 0);

        internal UdpMessageEventArgs(LocalUdp node, int target, int sender, int messageNumber, byte[] bytes)
            :base(node, target, sender, messageNumber, bytes)
        {
        }
    }

    #endregion

    #region UdpExceptionEventArgs

    public class UdpExceptionEventArgs : EventArgs
    {
        public readonly LocalUdp Udp;
        public readonly Exception Exception;
        public bool Handled = false;

        internal UdpExceptionEventArgs(LocalUdp node, Exception ex)
        {
            Udp = node;
            Exception = ex;
        }

        internal UdpExceptionEventArgs(LocalUdp node, string msg, Exception ex)
            : this(node, new Exception(msg, ex))
        {
        }

        internal UdpExceptionEventArgs(LocalUdp node, string msg)
            : this(node, new Exception(msg))
        {
        }

        public override string ToString()
            => Exception.ToString();
    }

    #endregion

    public class LocalUdp : CommunicationHandler<int>
    {
        protected static readonly byte[] EmptyByteArray = new byte[0];
        private const int DefaultBufferSize = 8192;
        public const int DefaultTimeout = int.MaxValue;

        public readonly int UdpPort;
        public static readonly int BroadcastNodeId = 0;
        private UdpClient udpClient;
        private IPAddress localAddress;
        private readonly IPEndPoint broadcastEndPoint;
        public int Timeout;
        public int BufferSize;

        public LocalUdp(int udpPort, Logger logger)
            : base(udpPort, logger)
        {
            UdpPort = udpPort;
            Timeout = DefaultTimeout;
            BufferSize = DefaultBufferSize;
            localAddress = IPAddress.Parse("127.0.0.1");
            broadcastEndPoint = new(IPAddress.Broadcast, UdpPort);

            // Open the port and start listening
            udpClient = new UdpClient();
            udpClient.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            var ip = new IPEndPoint(IPAddress.Any, UdpPort);
            udpClient.Client.Bind(ip);
            udpClient.Client.Setup(BufferSize, Timeout);
            udpClient.BeginReceive(StaticUdpReceiveCallback, this);
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

        internal UdpMessageEventArgs Parse(bool isUdp, byte[] raw)
            => Parse(raw ?? EmptyByteArray, 0, raw?.Length ?? 0);

        internal UdpMessageEventArgs Parse(byte[] raw, int offset, int length)
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

            return new UdpMessageEventArgs(this, target, sender, messageNumber, bytes);
        }

        #endregion

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
            udpClient?.Send(formattedMessage, formattedMessage.Length, broadcastEndPoint);
        }

        #endregion

        #region Incoming

        private static void StaticUdpReceiveCallback(IAsyncResult ar)
        {
            var udp = (LocalUdp) ar.AsyncState;
            if (udp.UdpReceiveCallback(ar))
                udp.udpClient.BeginReceive(StaticUdpReceiveCallback, udp);
        }

        private bool UdpReceiveCallback(IAsyncResult ar)
        {
            try
            {
                var from = new IPEndPoint(IPAddress.Broadcast, UdpPort);
                var bytes = udpClient?.EndReceive(ar, ref from);
                if (bytes == null)
                {
                    Log("Local UDP client closed");
                    return false;
                }

                var e = Parse(true, bytes);
                if (IsForThisNode(e))
                {
                    try
                    {
                        MessageReceived?.Invoke(this, e);
                        if (e.FormattedResponse != null)
                            udpClient.Send(e.FormattedResponse, e.FormattedResponse.Length, broadcastEndPoint);
                    }
                    catch (Exception ex)
                    {
                        Log(ex);
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex);
                return false;
            }

            return true; // to listen some more
        }

        private bool IsForThisNode(LocalUdpMessageEventArgs e)
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
