using Korn.Utils;
using System;
using System.Diagnostics;

namespace Korn.Com;
public unsafe class WMIProcessWatcher : IDisposable
{
    public WMIProcessWatcher()
    {
        context = new WMIContext();

        var sink = WbemObjectSinkImpl.Create();
        this.sink = sink;
        sink.AddRef();
        sinkImplementation = sink.GetImplementation();

        context.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'", *(void**)&sink);
    }

    WMIContext context;
    WbemObjectSinkImpl sink;
    WbemObjectSinkImpl.Implementation* sinkImplementation;

    public void SetProcessCreatedHandler(WMIProcessCreatedDelegate? handler) => sinkImplementation->SetProcessStartedHandler(handler);

    public void Dispose()
    {
        var sink = this.sink;
        context.CancelAsyncCall(*(void**)&sink);
        context.Dispose();        
        sink.Release();
    }
}