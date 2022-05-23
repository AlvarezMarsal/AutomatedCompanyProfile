using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public abstract class NamedPipeBase : BaseThreaded
    {
        public readonly string PipeName;
        public readonly Func<string, string> Handler;
        private readonly byte[] buffer;
        protected PipeStream stream;

        public NamedPipeBase(string pipeName, Func<string, string> handler, Logger logger) : base(pipeName, logger)
        {
            PipeName = pipeName;
            Handler = handler;
            buffer = new byte[4096];
        }

        public string Read()
        {
            try
            {
                var i = stream.ReadByte();
                if (i < 0)
                    return null;
                var len = i * 256;

                i = stream.ReadByte();
                if (i < 0)
                    return null;
                len += i;

                stream.Read(buffer, 0, len);
                return Encoding.UTF8.GetString(buffer, 0, len);
            }
            catch
            {
                return null;
            }

        }

        public void Write(string str)
        {
            if (str == null)
            {
                stream.Close();
                stream.Dispose();
                stream = null;
            }
            else if (str.Length == 0)
            {
                // do nothing
            }
            else
            {
                var len = str.Length;
                buffer[0] = (byte)(len / 256);
                buffer[1] = (byte)(len & 255);
                Encoding.UTF8.GetBytes(str, 0, str.Length, buffer, 2);
                stream.Write(buffer, 0, len + 2);
                stream.Flush();
            }
        }

        protected override void Stopping()
        {
            base.Stopping();

            stream.Close();
            stream.Dispose();
            stream = null;
        }
    }
}
