using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public static class ExtensionMethods
    {
        public static void SetupBeforeOpen(this Socket socket, int bufferSize = 0, int timeOut = int.MaxValue)
        {
            socket.ReceiveTimeout = timeOut;
            socket.SendTimeout = timeOut;

            bufferSize = Math.Max(8192, bufferSize);
            if (socket.ReceiveBufferSize < bufferSize)
                socket.ReceiveBufferSize = bufferSize;
            if (socket.SendBufferSize < bufferSize)
                socket.SendBufferSize = bufferSize;
        }

        public static void SetupAfterOpen(this Socket socket, int bufferSize = 0, int timeOut = int.MaxValue)
        {
            // try { socket.NoDelay = true; } catch { }
            socket.DontFragment = true;
        }

        public static void Setup(this Socket socket, int bufferSize = 0, int timeOut = int.MaxValue)
        {
            SetupBeforeOpen(socket, bufferSize, timeOut);
            SetupAfterOpen(socket, bufferSize, timeOut);
        }


    }
}
