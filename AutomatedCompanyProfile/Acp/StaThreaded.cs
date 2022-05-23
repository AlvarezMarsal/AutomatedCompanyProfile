using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public abstract class StaThreaded : BaseThread
    {
        public StaThreaded(string name, Log log) : base(name, log)
        {
        }

        protected override void Starting()
        {
            Thread.SetApartmentState(ApartmentState.STA);
            base.Starting();
        }

        override protected void RunWrapper()
        {
            MessageFilter.Register(); // STA threads have to pump messages
            base.RunWrapper();
        }

        // This is handled in the base class
        /*
        public override void Yield()
        {
            Thread.CurrentThread.Join(100);
            base.Yield();
        }
        */
    }
}
