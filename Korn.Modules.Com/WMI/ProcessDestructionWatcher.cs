using Korn.Modules.Com.Wmi.Internal;
using Korn.Modules.WinApi;
using Korn.Modules.WinApi.Ole;
using System.Runtime.InteropServices;

namespace Korn.Modules.Com.Wmi;
public unsafe class ProcessDestructionWatcher
{
    public ProcessDestructionWatcher(OleContext context)
    {
        this.context = context;

        implementation = Implementation.Create();
        sink = implementation->GetSink();
        sink.AddRef();

        context.ExecNotificationQuery("SELECT * FROM Win32_ProcessStopTrace", sink);
    }

    OleContext context;
    WbemObjectSink sink;
    Implementation* implementation;

    public void SetHandler(ProcessDestructedDelegate? handler) => implementation->SetProcessStartedHandler(handler);

    bool disposed;
    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;

        context.CancelAsyncCall(sink);
        sink.Release();
    }

    ~ProcessDestructionWatcher() => Dispose();

    struct Implementation
    {
        const uint WBEM_S_NO_ERROR = 0x00000000;
        const uint E_NOINTERFACE = 0x80004002;
        const ushort VT_NULL = 1;
        const ushort VT_EMPTY = 0;

        WbemObjectSinkImplementation baseClass;
        nint processStartHandler;

        public void SetProcessStartedHandler(ProcessDestructedDelegate? handler)
        {
            if (handler is null)
            {
                processStartHandler = default;
                return;
            }

            processStartHandler = Marshal.GetFunctionPointerForDelegate(handler);
        }

        void HandleProcessDestructed(DestructedProcess process)
        {
            if (processStartHandler == default)
                return;

            Marshal.GetDelegateForFunctionPointer<ProcessDestructedDelegate>(processStartHandler)(process);
        }

        public WbemObjectSink GetSink() => baseClass.GetSink();

        [UnmanagedCallersOnly]
        static uint Indicate(Implementation* self, int count, WbemClassObject* array)
        {
            int result;
            OleVariant property1;

            for (var index = 0; index < count; index++)
            {
                var process = new DestructedProcess();
                var obj = array[index];

                result = obj.Get("ProcessId", default, &property1, null, null);
                if (result >= 0 && property1.VariantType is not VT_EMPTY and not VT_NULL)
                    process.ID = property1.Int32;

                result = obj.Get("ParentProcessID", default, &property1, null, null);
                if (result >= 0 && property1.VariantType is not VT_EMPTY and not VT_NULL)
                    process.ParentID = property1.Int32;

                result = obj.Get("ProcessName", default, &property1, null, null);
                if (result >= 0 && property1.VariantType is not VT_EMPTY and not VT_NULL)
                    process.Name = ReadProcessNameWithoutExtension(property1.Bstr);
                else process.Name = string.Empty;
                Ole32.VariantClear(&property1);

                self->HandleProcessDestructed(process);
            }

            return WBEM_S_NO_ERROR;

            string ReadProcessNameWithoutExtension(char* bstr)
            {
                const ulong ExeSuffix = '.' << 0x00 | 'e' << 0x10 | (ulong)'x' << 0x20 | (ulong)'e' << 0x30;

                var len = ((int*)bstr)[-1] / 2;
                if (len > 4)
                    if (*(ulong*)&bstr[len - 4] == ExeSuffix)
                        len -= 4;

                return new string(bstr, 0, len);
            }
        }

        public static Implementation* Create()
        {
            var implementation = (Implementation*)Marshal.AllocCoTaskMem(sizeof(Implementation));
            implementation->baseClass.Initialize((delegate* unmanaged<Implementation*, int, WbemClassObject*, uint>)&Indicate);
            implementation->processStartHandler = default;

            return implementation;
        }
    }
}