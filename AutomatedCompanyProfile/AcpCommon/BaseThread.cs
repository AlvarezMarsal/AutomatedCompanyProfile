using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public abstract class BaseThread : LogUser, IDisposable
    {
        public Exception Exception;
        private string name;
        private bool disposed;
        protected Thread Thread { get; private set; }
        private readonly ManualResetEvent stopSignal = new ManualResetEvent(false);     // signal TO this thread that it should stop
        private readonly ManualResetEvent stoppedSignal = new ManualResetEvent(false);  // signal that the thread HAS stopped
        private readonly ManualResetEvent startedSignal = new ManualResetEvent(false);  // signal that the thread HAS stopped
        private readonly ConcurrentDictionary<string, object> tags = new ConcurrentDictionary<string, object>();

        public System.Threading.ThreadState State => Thread?.ThreadState ?? System.Threading.ThreadState.Unstarted;

        protected BaseThread(string name, LogRecorder recorder) : base(recorder)
        {
            Name = name;
        }


        public string Name
        {
            get => name;
            set
            {
                if (name != null)
                    Log("Changing name of thread to " + value);
                name = value;
                base.ThreadName = name;
            }
        }


        public void SetTag(string name, object o)
        {
            var old = tags.GetOrAdd(name, o);
            if (!old.Equals(o) && (old is IDisposable d))
                d.Dispose();
        }

        public object GetTag(string name, object o)
        {
            return tags.TryGetValue(name, out var old) ? old : null;
        }

        public bool GetTag<T>(string name, out T value)
            => GetTag(name, out value, default);

        public bool GetTag<T>(string name, out T value, T defaultValue = default)
        {
            if (tags.TryGetValue(name, out var raw) && (raw is T t))
            {
                value = t;
                return true;
            }
            value = defaultValue;
            return false;
        }

        /*
        public bool IsRunning
        {
            get
            {
                var s = State & (System.Threading.ThreadState.Aborted | System.Threading.ThreadState.Stopped);
                return (s == 0);
            }
        }
        */

        // overrides must call base.Start after their own logic!
        // Remember that this is executed on some other thread
        public bool Start(int wait = Timeout.Infinite)
        {
            if (State != System.Threading.ThreadState.Unstarted)
            {
                Warning("Call to start " + Name + " while already running");
                return false;
            }

            stopSignal.Reset();
            stoppedSignal.Reset();
            startedSignal.Reset();

            Thread = new Thread(() => RunWrapper());
            Thread.Name = Name;
            try
            {
                Starting(); // starting runs on the CALLER's thread!
            }
            catch (Exception ex)
            {
                Exception = ex;
                return false;
            }
 
            Thread.Start();
            var signals = new WaitHandle[] { stoppedSignal, startedSignal };
            var result = WaitHandle.WaitAny(signals, wait);
            return (result == 1); // true if it is running
        }

        // Can be called (indirectly) from any thread
        public void Yield(int milliseconds = 100) 
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                Thread.Join(milliseconds);
            else
                stopSignal.WaitOne(milliseconds);
        }

        // overrides should call base.Stop() before their own logic!
        // This might run on any thread!
        public bool Stop(int wait = Timeout.Infinite)
        {
            stopSignal.Set();
            if (Thread.CurrentThread.ManagedThreadId == Thread.ManagedThreadId)
                return false; // a thread can't wait on its own stoppage
            return stoppedSignal.WaitOne(wait);
        }

        public void Abort(int wait = Timeout.Infinite)
        {
            Thread.Abort();
            if (Thread.CurrentThread.ManagedThreadId == Thread.ManagedThreadId)
                return; // a thread can't wait on its own stoppage
            stoppedSignal.WaitOne(wait);
        }

        public bool IsStopped => startedSignal.WaitOne(0) && stoppedSignal.WaitOne(0);

        protected virtual void RunWrapper()
        {
            try
            {
                Started();
            }
            catch (ThreadAbortException)
            {
                stopSignal.Set();
            }
            catch (Exception ex)
            {
                Exception = ex;
                stopSignal.Set();
            }

            if (!stopSignal.WaitOne(0))
            {
                startedSignal.Set();
                try
                {
                    Run();
                }
                catch (ThreadAbortException)
                {
                    stopSignal.Set();
                }
                catch (Exception ex)
                {
                    Exception = ex;
                    stopSignal.Set();
                }
                finally
                {
                    Thread.BeginCriticalRegion();
                    try
                    {
                        Stopping();
                        Stopped();
                    }
                    finally
                    {
                        stoppedSignal.Set();
                    }
                    Thread.EndCriticalRegion();
                }
            }
        }

        protected virtual void Starting() { }   // runs on caller's thread
        protected virtual void Started() { }    // runs on this.Thread
        protected virtual void Stopping() { }   // runs on this.Thread

        protected virtual void Stopped()        // runs on this.Thread
        {
            var values = tags.Values.ToArray();
            tags.Clear();
            foreach (var v in values)
            {
                if (v is IDisposable d)
                    d.Dispose();
            }
        }    

        protected bool ShouldKeepRunning => !ShouldStop;
        protected bool ShouldStop => stopSignal.WaitOne(0);
        //protected bool StopSignalReceived(int milliseconds = 0) => stopSignal.WaitOne(milliseconds);
        protected bool StopSignalReceived => stopSignal.WaitOne(0);

        // It is VERY important that the user's Run() method
        // sets IsRunning = true at some early point in its
        //  processing.
        protected abstract void Run();

        #region Dispose

        protected override void Dispose(bool disposing)
        {
            if (!disposed)
            {
                disposed = true;
                if (disposing)
                {
                    try
                    {
                        Stop(0);
                    }
                    catch (Exception ex)
                    {
                        Log(ex);
                    }
                }
            }
            base.Dispose(disposing);
        }

        #endregion

        public void ZapComObject<T>(ref T o) where T : class
        {
            if (o != null)
            {
                while (Marshal.FinalReleaseComObject(o) > 0)
                    ;
                o = null;
            }
        }

        public void ReleaseComObject(object o)
        {
            int refcnt = Marshal.FinalReleaseComObject(o);
            Debug.Assert(refcnt == 0);
        }

        public bool TryForPeriod(Func<bool> func, int milliseconds = 15 * 60 * 1000, int yieldMilliseconds = 250)
        {
            var giveUp = DateTime.Now.AddMilliseconds(milliseconds);
            while (true)
            {
                if (func())
                    return true;
                if (DateTime.Now > giveUp)
                    return false;
                Yield(yieldMilliseconds);
            }
        }

    }
}
