using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Acp
{
    public class NamedPipeClient : NamedPipeBase
    {
        public NamedPipeClient(string pipeName, Func<string, string> handler, Logger logger) 
            : base(pipeName, handler, logger)
        {
        }

        protected override void Run()
        {
            var s = new NamedPipeClientStream(
                ".",
                PipeName,
                PipeDirection.InOut, PipeOptions.None,
                TokenImpersonationLevel.Impersonation);

            stream = s;
            s.Connect();

            try
            {
                while (ShouldKeepRunning)
                {
                    var request = Read();
                    if (request == null)
                        break;
                    var response = Handler(request);
                    if (response == null)
                        break;
                    if (response.Length > 0)
                        Write(response);
                }
            }
            catch
            {
            }
        }

    }
}
