using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    #region MessageEventArgs

    public class MessageEventArgs<A> : EventArgs  // A is an address 
    {
        public readonly CommunicationHandler<A> CommunicationHandler;
        public virtual bool IsBroadcast => false;
        public readonly A Target;
        public readonly A Sender;
        public readonly byte[] MessageBytes;
        public readonly int MessageNumber;
        private string[] message;
        public string[] Message => message ??= ParseMessageBytes();
        internal byte[] FormattedResponse;

        internal MessageEventArgs(CommunicationHandler<A> node, A target, A sender, int messageNumber, byte[] bytes)
        {
            CommunicationHandler = node;
            Target = target;
            Sender = sender;
            MessageNumber = messageNumber;
            MessageBytes = bytes;
        }

        public void Respond(byte[] bytes)
            => Respond(MessageNumber + 1000000, bytes);

        public void Respond(byte[] bytes, int index, int count)
            => Respond(MessageNumber + 1000000, bytes, index, count);

        public void Respond(string first, string second = null, string third = null, string fourth = null)
            => Respond(MessageNumber + 1000000, first, second, third, fourth);

        public void Respond(int messageNumber, byte[] bytes)
            => FormattedResponse = CommunicationHandler.FormatMessage(Sender, messageNumber, bytes);

        public void Respond(int messageNumber, byte[] bytes, int index, int count)
            => FormattedResponse = CommunicationHandler.FormatMessage(Sender, messageNumber, bytes, index, count);

        public void Respond(int messageNumber, string first, string second = null, string third = null, string fourth = null)
            => FormattedResponse = CommunicationHandler.FormatMessage(Sender, messageNumber, first, second, third, fourth);

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
                var e = Udp.Parse(FormattedResponse);
                if (e != null)
                    b.AppendLine().Append("    " + e.ToString());
            }
            return b.ToString();
        }

    }

    #endregion MessageEventArgs

    #region ExceptionEventArgs

    public class ExceptionEventArgs<A> : EventArgs
    {
        public readonly CommunicationHandler<A> CommunicationHandler;
        public readonly Exception Exception;
        public bool Handled = false;

        internal ExceptionEventArgs(CommunicationHandler<A> node, Exception ex)
        {
            CommunicationHandler = node;
            Exception = ex;
        }

        internal ExceptionEventArgs(CommunicationHandler<A> node, string msg, Exception ex)
            : this(node, new Exception(msg, ex))
        {
        }

        internal ExceptionEventArgs(CommunicationHandler<A> node, string msg)
            : this(node, new Exception(msg))
        {
        }

        public override string ToString()
            => Exception.ToString();
    }

    #endregion


    public abstract class CommunicationHandler : LogUser, IDisposable
    {
        private static int nextMessageNumber = -1;

        public int NextMessageNumber => Interlocked.Increment(ref nextMessageNumber);

        protected CommunicationHandler(Logger logger) : base(logger)
        {
        }

        protected virtual void SendBytes(byte[] bytes)
            => SendBytes(bytes, 0, bytes.Length);
        protected abstract void SendBytes(byte[] bytes, int offset, int count);

        #region Buffer stuff

        protected int WriteIntoBuffer(byte[] message, int position, byte[] bytes, int offset, int count)
        {
            Array.Copy(bytes, offset, message, position, count);
            return position + count;
        }

        protected int WriteIntoBuffer(byte[] message, int position, byte[] bytes)
            => WriteIntoBuffer(message, position, bytes, 0, bytes.Length);

        protected int WriteIntoBuffer(byte[] message, int position, ushort u)
            => WriteIntoBuffer(message, position, BitConverter.GetBytes(u));

        protected int WriteIntoBuffer(byte[] message, int position, int i)
             => WriteIntoBuffer(message, position, BitConverter.GetBytes(i));

        protected int ReadFromBuffer(byte[] message, int position, int count, out byte[] bytes)
        {
            bytes = new byte[count];
            Array.Copy(message, position, bytes, 0, count);
            return position + count;
        }

        protected int ReadFromBuffer(byte[] message, int position, out ushort u)
        {
            u = BitConverter.ToUInt16(message, position);
            return position + sizeof(ushort);
        }

        protected int WriteIntoBuffer(byte[] message, int position, out int i)
        {
            i = BitConverter.ToInt32(message, position);
            return position + sizeof(int);
        }

        #endregion Buffer stuff

        #region IDispose

        private bool disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
                if (disposing)
                {
                    DisposeManagedObjects();
                }

                DisposeUnmanagedObjects();
                base.Dispose(disposing);
            }
        }

        protected virtual void DisposeManagedObjects()
        {
        }

        protected virtual void DisposeUnmanagedObjects()
        {
        }

        #endregion IDispose
    }

    public abstract class CommunicationHandler<S> : CommunicationHandler // S is the type of an address
    {
        public event EventHandler<MessageEventArgs<S>> MessageReceived;
        public event EventHandler<ExceptionEventArgs<S>> ExceptionOccurred;

        public S NodeId { get; protected set; }
        protected abstract int SizeofNodeId { get; }

        protected CommunicationHandler(Logger logger) : base(logger)
        {
        }

        protected CommunicationHandler(S nodeId, Logger logger) : base(logger)
        {
            NodeId = nodeId;
        }

        public virtual byte[] FormatMessage(S target, int messageNumber, byte[] bytes)
            => FormatMessage(target, messageNumber, bytes, 0, bytes?.Length ?? throw new ArgumentNullException("bytes"));
        
        public virtual byte[] FormatMessage(S target, int messageNumber, byte[] bytes, int offset, int count)
        {
            ushort ucount = (ushort) count;
            ushort totalLength = (ushort)(ucount + 2 * sizeof(short) + 2 * SizeofNodeId + sizeof(int));
            var message = new byte[totalLength];
            int o = WriteIntoBuffer(message, 0, totalLength);
            o = WriteIntoBuffer(message, o, ucount); // not really necessary -- consider removing
            o = WriteIntoBuffer(message, o, target);
            o = WriteIntoBuffer(message, o, NodeId);
            if (messageNumber < 0)
                messageNumber = NextMessageNumber;
            o = WriteIntoBuffer(message, o, messageNumber);
            WriteIntoBuffer(message, o, bytes, 0, count);
            return message;
        }

        protected abstract int WriteIntoBuffer(byte[] message, int position, S nodeId);

        public virtual void Send(S target, int messageNumber, byte[] bytes)
        {
            var message = FormatMessage(target, messageNumber, bytes, 0, bytes.Length);
            SendBytes(message);
        }

        public virtual void Send(S target, byte[] bytes)
            => Send(target, NextMessageNumber, bytes);
    }
}
