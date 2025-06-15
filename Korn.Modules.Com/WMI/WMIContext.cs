using Korn.Utils;
using System;

namespace Korn.Com;
public unsafe class WMIContext : IDisposable
{
    const uint WBEM_FLAG_SEND_STATUS = 0x80;

    public WMIContext()
    {
        int result;

        result = Ole32.CoInitializeEx(CoInit.Multithreaded);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to co initialize. result: {result}");

        result = Ole32.CoInitializeSecurity(null, CoAuthSvc.Register, null, RpcAuthnLevel.Default, RpcImplLevel.Impersonate, null, EoleAuthenticationCapabilities.None);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to co initialize security. result: {result}");

        WbemLocator locator;
        result = Ole32.CoCreateInstance(WbemLocator.CLSID.GUID, null, CoClassContext.InProcessServer, WbemLocator.IID.GUID, &locator);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to co create intance. result: {result}");
        this.locator = locator;

        WbemServices services;
        result = locator.ConnectServer("root\\CIMV2", null, null, null, default, null, default, &services);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to connect to server. result: {result}");
        this.services = services;

        result = Ole32.CoSetProxyBlanket(*(void**)&services, RpcAuthnSvc.WinNT, RpcAuthzSvc.None, null, RpcAuthnLevel.Call, RpcImplLevel.Impersonate, default, EoleAuthenticationCapabilities.None);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to set proxy blanket. result: {result}");
    }

    WbemLocator locator;
    WbemServices services;

    public void ExecNotificationQuery(string query, WbemObjectSink.Interface sink)
    {
        var queryBstr = Ole32.SysAllocString(query);
        var result = services.ExecNotificationQueryAsync("WQL", queryBstr, WBEM_FLAG_SEND_STATUS, default, sink);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to execute notification query. result: {result}");
    }

    public void CancelAsyncCall(WbemObjectSink.Interface sink)
    {
        var result = services.CancelAsyncCall(sink);
        if (result < 0)
            throw new Exception($"Korn.Com.WMIProcessWatcher: Unable to cancel async call when disposing watcher. result: {result}");
    }

    bool disposed;
    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;

        locator.Release();
        services.Release();
    }

    ~WMIContext() => Dispose();
}