using Korn.Com.Wmi.Internal;
using System;

namespace Korn.Com.Wmi;
public class ProcessWatcher : IDisposable
{
    public ProcessWatcher()
    {
        context = new OleContext();
        creationWatcher = new ProcessCreationWatcher(context);
        creationWatcher.SetHandler(OnProcessCreated);

        destructionWatcher = new ProcessDestructionWatcher(context);
        destructionWatcher.SetHandler(OnProcessDestructed);
    }

    OleContext context;
    ProcessCreationWatcher creationWatcher;
    ProcessDestructionWatcher destructionWatcher;

    public Action<CreatedProcess>? ProcessStarted;
    public Action<DestructredProcess>? ProcessStopped;

    void OnProcessCreated(CreatedProcess process) => ProcessStarted?.Invoke(process);
    void OnProcessDestructed(DestructredProcess process) => ProcessStopped?.Invoke(process);

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