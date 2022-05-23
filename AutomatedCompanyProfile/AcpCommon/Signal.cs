using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Acp
{
    public class Signal
    {
        public readonly string Name;
        internal readonly EventWaitHandle Handle;
        public const EventResetMode Mode = EventResetMode.ManualReset;
        public readonly bool IsInterprocess;

        public Signal(string name) : this(name, false) { }
        public Signal(string name, bool interprocess)
        {
            Name = name;
            IsInterprocess = interprocess;
            if (IsInterprocess)
            {
                if (EventWaitHandle.TryOpenExisting(name, out var handle))
                    Handle = handle;
                else
                    Handle = new EventWaitHandle(false, Mode, name);
            }
            else
            {
                Handle = new EventWaitHandle(false, Mode);
            }
        }

        internal Signal(string name, EventWaitHandle handle)
        {
            Name = name;
            Handle = handle;
        }

        public void WaitForever() => Handle.WaitOne();
        public bool Wait(int milliseconds) => Handle.WaitOne(milliseconds);
        public bool Wait(TimeSpan ts) => Handle.WaitOne(ts);
        public void Set() => Handle.Set();
        public void Reset() => Handle.Reset();
        public bool State
        {
            get => Handle.WaitOne(0);
            set { if (value) Handle.Set(); else Handle.Reset(); }
        }

        public bool IsSet => Handle.WaitOne(0);
        public void SetAndWait(WaitHandle toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn);
        public void SetAndWait(Signal toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn.Handle);
        public void SetAndWait(int milliseconds, WaitHandle toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn, milliseconds, false);
        public void SetAndWait(int milliseconds, Signal toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn.Handle, milliseconds, false);
        public void SetAndWait(TimeSpan ts, WaitHandle toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn, ts, false);
        public void SetAndWait(TimeSpan ts, Signal toWaitOn) => WaitHandle.SignalAndWait(Handle, toWaitOn.Handle, ts, false);
    }

    public class SignalCollection
    {
        private ConcurrentDictionary<string, Signal> signals;

        public Signal Get(string name, bool interprocess)
        {
            if (signals == null)
                signals = new ConcurrentDictionary<string, Signal>();

            Signal signal;
            while (true) // theortically, we could need to do this over and over
            {
                if (signals.TryGetValue(name, out signal))
                {
                    if (interprocess != signal.IsInterprocess)
                        throw new AcpException("A signal has already been created with that name, but different characteristics");
                    break;
                }

                signal = new Signal(name, interprocess);
                if (signals.TryAdd(name, signal))
                    break;
            }

            return signal;
        }
            
        public Signal Get(string name)
        {
            Signal signal = null;
            if ((signals != null) && signals.TryGetValue(name, out signal))
                return signal;

            if (EventWaitHandle.TryOpenExisting(name, out var handle))
            {
                if (signals == null)
                    signals = new ConcurrentDictionary<string, Signal>();
                signal = new Signal(name, handle);

                while (true)
                {
                    if (signals.TryAdd(name, signal))
                        return signal;
                    if (signals.TryGetValue(name, out var s))
                        return s;
                }
            }

            return signal;
        }

        public Signal this[string name]
        {
            get => Get(name);
        }

        public void WaitForever(string name) => SignalDo(name, s => s.WaitForever());
        public bool Wait(string name, int milliseconds) => SignalDo(name, (s) => s.Wait(milliseconds));
        public bool Wait(string name, TimeSpan ts) => SignalDo(name, (s) => s.Wait(ts));
        public void Set(string name) => SignalDo(name, s => s.Set());
        public void Reset(string name) => SignalDo(name, s => s.Reset());
        public bool IsSet(string name) => SignalDo(name, s => s.IsSet);

        private void SignalDo(string name, Action<Signal> action)
        {
            var signal = this[name] ?? throw new AcpException("There is no signal with that name");
            action(signal);
        }

        private T SignalDo<T>(string name, Func<Signal, T> action)
        {
            var signal = this[name] ?? throw new AcpException("There is no signal with that name");
            return action(signal);
        }
    }
}
