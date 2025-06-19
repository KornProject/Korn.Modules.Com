using Korn.Modules.Com.Wmi.Internal;
using System;

namespace Korn.Modules.Com.Wmi;
public class ProcessWatcher : IDisposable
{
    public ProcessWatcher()
    {
        creationDelegate = OnProcessCreated;
        destructionDelegate = OnProcessDestructed;

        context = new OleContext();
        creationWatcher = new ProcessCreationWatcher(context);
        creationWatcher.SetHandler(creationDelegate);

        destructionWatcher = new ProcessDestructionWatcher(context);
        destructionWatcher.SetHandler(destructionDelegate);
    }

    ProcessCreatedDelegate creationDelegate;
    ProcessDestructedDelegate destructionDelegate;

    OleContext context;
    ProcessCreationWatcher creationWatcher;
    ProcessDestructionWatcher destructionWatcher;

    public Action<CreatedProcess>? ProcessStarted;
    public Action<DestructedProcess>? ProcessStopped;

    void OnProcessCreated(CreatedProcess process) => ProcessStarted?.Invoke(process);
    void OnProcessDestructed(DestructedProcess process) => ProcessStopped?.Invoke(process);

    bool disposed;
    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;

        creationWatcher.Dispose();
        destructionWatcher.Dispose();
        context.Dispose();
    }

    ~ProcessWatcher() => Dispose();
}