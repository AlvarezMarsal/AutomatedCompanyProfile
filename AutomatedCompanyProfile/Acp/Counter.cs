using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    class Counter
    {
        private long next = 0;
        [ThreadStatic] private static long mostRecent; // do not initialize

        public Counter(long n = 0) { next = n; }

        public long Next
        {
            get
            {
                var n = Interlocked.Increment(ref next);
                mostRecent = n;
                return n;
            }
        }

        public long MostRecent // The most recent on the current thread
        {
            get
            {
                return mostRecent;
            }
        }

    }
}
