using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class NamedPipeServer : NamedPipeBase
    {
        public NamedPipeServer(string name, Func<string, string> handler, Logger logger) 
            : base(name, handler, logger)
        {
        }

        protected override void Run()
        {
            var s = new NamedPipeServerStream(Name, PipeDirection.InOut);
            stream = s;

            try
            {
                string request = null;

                while (ShouldKeepRunning)
                {
                    if (!s.IsConnected)
                    {
                        request = null;
                        s.WaitForConnection();
                    }
                    var response = Handler(request);
                    if (response == null)
                        break;
                    Write(response);
                    request = Read();
                }
            }
            catch
            {
            }
        }
    }
}
