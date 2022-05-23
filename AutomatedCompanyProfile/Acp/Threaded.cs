using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public abstract class Threaded : BaseThread
    {
        public Threaded(string name, Log log) : base(name, log)
        {
        }
        public Threaded(string name, LogRecorder recorder) : base(name, recorder)
        {
        }
    }
}
