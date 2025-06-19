using Korn.Modules.Com.Wmi.Internal;
using Korn.Modules.WinApi;
using Korn.Modules.WinApi.Ole;
using System;
using System.Runtime.InteropServices;

namespace Korn.Modules.Com.Wmi;
public unsafe class ProcessCreationWatcher : IDisposable
{
    public ProcessCreationWatcher(OleContext context)
    {
        this.context = context;

        implementation = Implementation.Create();
        sink = implementation->GetSink();
        sink.AddRef();

        context.ExecNotificationQuery("SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'", sink);
    }

    OleContext context;
    WbemObjectSink sink;
    Implementation* implementation;

    public void SetHandler(ProcessCreatedDelegate? handler) => implementation->SetProcessStartedHandler(handler);

    bool disposed;
    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;

        context.CancelAsyncCall(sink);
        sink.Release();
    }

    ~ProcessCreationWatcher() => Dispose();

    struct Implementation 
    {
        const uint WBEM_S_NO_ERROR = 0x00000000;
        const uint E_NOINTERFACE = 0x80004002;
        const ushort VT_NULL = 1;
        const ushort VT_EMPTY = 0;

        WbemObjectSinkImplementation baseClass;
        nint processStartHandler;

        public void SetProcessStartedHandler(ProcessCreatedDelegate? handler)
        {
            if (handler is null)
            {
                processStartHandler = default;
                return;
            }

            processStartHandler = Marshal.GetFunctionPointerForDelegate(handler);
        }

        void HandleProcessStarted(CreatedProcess process)
        {
            if (processStartHandler == default)
                return;

            Marshal.GetDelegateForFunctionPointer<ProcessCreatedDelegate>(processStartHandler)(process);
        }

        public WbemObjectSink GetSink() => baseClass.GetSink();

        [UnmanagedCallersOnly]
        static uint Indicate(Implementation* self, int count, WbemClassObject* array)
        {
            int result;
            OleVariant property1, property2;

            for (var index = 0; index < count; index++)
            {
                var process = new CreatedProcess();
                var obj = array[index];

                result = obj.Get("TargetInstance", default, &property1, default, default);
                if (result >= 0)
                {
                    IUnknown procedure = *(IUnknown*)&property1.Unknown;
                    result = procedure.QueryInterface<WbemClassObject>(&obj);
                    if (result >= 0)
                    {
                        result = obj.Get("ProcessId", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.ID = property2.Int32;

                        result = obj.Get("ParentProcessId", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.ParentID = property2.Int32;

                        result = obj.Get("CreationDate", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.CreationTime = DMTFDateConvertor.ConvertToTicks(property2.Bstr);
                        else process.CreationTime = default;
                        Ole32.VariantClear(&property2);

                        result = obj.Get("Name", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.Name = ReadProcessNameWithoutExtension(property2.Bstr);
                        else process.Name = string.Empty;
                        Ole32.VariantClear(&property2);

                        result = obj.Get("CommandLine", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.CommandLine = Marshal.PtrToStringBSTR((nint)property2.Bstr);
                        else process.CommandLine = string.Empty;
                        Ole32.VariantClear(&property2);

                        result = obj.Get("ExecutablePath", default, &property2, null, null);
                        if (result >= 0 && property2.VariantType is not VT_EMPTY and not VT_NULL)
                            process.ExecutablePath = Marshal.PtrToStringBSTR((nint)property2.Bstr);
                        else process.ExecutablePath = string.Empty;
                        Ole32.VariantClear(&property2);
                    }
                }
                Ole32.VariantClear(&property1);

                self->HandleProcessStarted(process);
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