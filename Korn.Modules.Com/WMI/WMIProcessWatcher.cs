using System;

namespace Korn.Com;
public unsafe class WMIProcessWatcher : IDisposable
{
    public WMIProcessWatcher()
    {
        context = new WMIContext();

        sink = WbemObjectSinkImpl.Create();
        sink.AddRef();
        sinkImplementation = sink.GetImplementation();

        context.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'", sink);
    }

    WMIContext context;
    WbemObjectSinkImpl sink;
    WbemObjectSinkImpl.Implementation* sinkImplementation;

    public void SetProcessCreatedHandler(WMIProcessCreatedDelegate? handler) => sinkImplementation->SetProcessStartedHandler(handler);

    bool disposed;
    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;

        context.CancelAsyncCall(sink);
        sink.Release();
        context.Dispose();        
    }

    ~WMIProcessWatcher() => Dispose();
}