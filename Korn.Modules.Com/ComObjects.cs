using Korn.Utils;
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Serialization;
using static System.Runtime.InteropServices.LayoutKind;

#pragma warning disable CS8500 // This takes the address of, gets the size of, or declares a pointer to a managed type
namespace Korn.Com;
public unsafe static class Extensions
{
    public static T As<T>(this IUnknown.Interface self) => *(*(T**)&self + 1);

    public static int QueryInterface(this IUnknown.Interface self, IID iid, void* obj)
        => ((delegate* unmanaged<IUnknown, IID*, void*, int>)self[0])(self.As<IUnknown>(), &iid, obj);

    public static int QueryInterface<T>(this IUnknown.Interface self, void* obj) where T : IUnknown.Interface
        => self.QueryInterface(T.IID, obj);

    public static int AddRef(this IUnknown.Interface self)
    => ((delegate* unmanaged<IUnknown, int>)self[1])(self.As<IUnknown>());

    public static int Release(this IUnknown.Interface self)
        => ((delegate* unmanaged<IUnknown, int>)self[2])(self.As<IUnknown>());

    public static int ConnectServer(
        this WbemLocator.Interface self,
        ReadOnlySpan<char> networkResource,
        ReadOnlySpan<char> user,
        ReadOnlySpan<char> password,
        ReadOnlySpan<char> locale,
        int securityFlags,
        ReadOnlySpan<char> authority,
        WbemContext context,
        WbemServices* pNamespace)
    {
        fixed (char* networkResourcePointer = networkResource,
                     userPointer = user,
                     passwordPointer = password,
                     localePointer = locale,
                     authorityPointer = authority)
            return ((delegate* unmanaged<WbemLocator, char*, char*, char*, char*, int, char*, WbemContext, WbemServices*, int>)self[3])(
                self.As<WbemLocator>(),
                networkResourcePointer,
                userPointer,
                passwordPointer,
                localePointer,
                securityFlags,
                authorityPointer,
                context,
                pNamespace
            );
    }

    public static int ExecNotificationQueryAsync(
        this WbemServices.Interface self,
        ReadOnlySpan<char> queryLanguage,
        char* query,
        uint flags,
        WbemContext context,
        void* implementation)
    {
        fixed (char* queryLanguagePointer = queryLanguage)
            return ((delegate* unmanaged<WbemServices, char*, char*, uint, WbemContext, void*, int>)self[23])(
                self.As<WbemServices>(),
                queryLanguagePointer,
                query,
                flags,
                context,
                implementation
            );
    }

    public static int CancelAsyncCall( this WbemServices.Interface self, void* implementation) 
        => ((delegate* unmanaged<WbemServices, void*, int>)self[4])(self.As<WbemServices>(), implementation);

    public static int Get(this WbemClassObject.Interface self, ReadOnlySpan<char> name, int flags, OleVariant* variant, int* type, int* flavor)
    {
        fixed (char* namePointer = name)
            return ((delegate* unmanaged<WbemClassObject, char*, int, OleVariant*, int*, int*, int>)self[4])(
                self.As<WbemClassObject>(),
                namePointer,
                flags,
                variant,
                type,
                flavor
            );
    }

    public static WbemObjectSinkImpl.Implementation* GetImplementation(this WbemObjectSinkImpl.Interface self)
        => ((delegate* unmanaged<WbemObjectSinkImpl, WbemObjectSinkImpl.Implementation*>)self[5])(self.As<WbemObjectSinkImpl>());
}


[StructLayout(Sequential, Size = 8)]
public unsafe struct IUnknown : IUnknown.Interface
{
    public static IID IID { get; private set; } = "0000000000000000c000000000000046"u8;
    public interface Interface
    {
        public static abstract IID IID { get; }

        void** vtable => *(void***)this.As<nint>();

        void* this[int index] => vtable[index];
    }
}

[StructLayout(Sequential, Size = 8)]
public struct WbemLocator : WbemLocator.Interface
{
    public static IID IID { get; private set; } = "dc12a687737f11cf884d00aa004b2e24"u8;
    public static IID CLSID { get; private set; } = "4590f8111d3a11d0891f00aa004b2e24"u8;
    public interface Interface : IUnknown.Interface;
}

[StructLayout(Sequential, Size = 8)]
public struct WbemContext : WbemContext.Interface
{
    public static IID IID { get; private set; } = "44aca674e8fc11d0a07c00c04fb68820"u8;
    public interface Interface : IUnknown.Interface;
}

[StructLayout(Sequential, Size = 8)]
public struct WbemServices : WbemServices.Interface
{
    public static IID IID { get; private set; } = "9556dc99828c11cfa37e00aa003240c7"u8;
    public interface Interface : IUnknown.Interface;
}

[StructLayout(Sequential, Size = 8)]
public struct WbemObjectSink : WbemObjectSink.Interface
{
    public static IID IID { get; private set; } = "7c857801738111cf884d00aa004b2e24"u8;
    public interface Interface : IUnknown.Interface;
}

[StructLayout(Sequential, Size = 8)]
public struct WbemClassObject : WbemClassObject.Interface
{
    public static IID IID { get; private set; } = "dc12a681737f11cf884d00aa004b2e24"u8;
    public interface Interface : IUnknown.Interface;
}

[StructLayout(Sequential, Size = 8)]
public unsafe struct WbemObjectSinkImpl : WbemObjectSinkImpl.Interface
{
    public static IID IID { get; private set; } = default;
    public interface Interface : WbemObjectSink.Interface;

    public static WbemObjectSinkImpl Create()
    {
        var implementation = Implementation.Create();
        return *(WbemObjectSinkImpl*)&implementation;
    }

    public unsafe struct Implementation
    {
        const uint WBEM_S_NO_ERROR = 0x00000000;
        const uint E_NOINTERFACE = 0x80004002;
        const ushort VT_NULL = 1;
        const ushort VT_EMPTY = 0;

        static Implementation()
        {
            var vtable = vtableDef = (void**)Marshal.AllocCoTaskMem(sizeof(void*) * 6);
            *vtable++ = (delegate* unmanaged<Implementation*, IID*, void**, uint>)&QueryInterface;
            *vtable++ = (delegate* unmanaged<Implementation*, uint>)&AddRef;
            *vtable++ = (delegate* unmanaged<Implementation*, uint>)&Release;
            *vtable++ = (delegate* unmanaged<Implementation*, int, WbemClassObject*, uint>)&Indicate;
            *vtable++ = (delegate* unmanaged<Implementation*, int, char*, WbemClassObject, uint>)&SetStatus;
            *vtable++ = (delegate* unmanaged<Implementation*, Implementation*>)&GetImplementation;
        }
        static void** vtableDef;

        void** vtable;
        uint refs;
        nint processStartHandler;

        public void SetProcessStartedHandler(WMIProcessCreatedDelegate? handler)
        {
            if (handler is null)
            {
                processStartHandler = default;
                return;
            }

            processStartHandler = Marshal.GetFunctionPointerForDelegate(handler);
        }

        void HandleProcessStarted(WMIProcessIntance process)
        {
            if (processStartHandler == default)
                return;

            Marshal.GetDelegateForFunctionPointer<WMIProcessCreatedDelegate>(processStartHandler)(process);
        }

        void Delete() => Marshal.FreeCoTaskMem((nint)Unsafe.AsPointer(ref this));

        [UnmanagedCallersOnly]
        static uint QueryInterface(Implementation* self, IID* iid, void** ppv)
        {
            if (*iid == IUnknown.IID || *iid == WbemObjectSink.IID)
            {
                *ppv = self;
                return WBEM_S_NO_ERROR;
            }

            return E_NOINTERFACE;
        }

        [UnmanagedCallersOnly]
        static uint AddRef(Implementation* self) => Interlocked.Increment(ref self->refs);

        [UnmanagedCallersOnly]
        static uint Release(Implementation* self)
        {
            var refs = Interlocked.Decrement(ref self->refs);

            if (refs == 0)
                self->Delete();

            return refs;
        }

        [UnmanagedCallersOnly]
        static uint Indicate(Implementation* self, int count, WbemClassObject* array)
        {
            int result;
            OleVariant property1, property2;

            for (var index = 0; index < count; index++)
            {
                var process = new WMIProcessIntance();
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

        [UnmanagedCallersOnly]
        static uint SetStatus(Implementation* self, int flags, char* param, WbemClassObject array) => WBEM_S_NO_ERROR;

        [UnmanagedCallersOnly]
        static Implementation* GetImplementation(Implementation* self) => self;

        public static Implementation* Create()
        {
            var handler = (Implementation*)Marshal.AllocCoTaskMem(sizeof(Implementation));
            handler->refs = 0;
            handler->vtable = vtableDef;
            handler->processStartHandler = default;

            return handler;
        }
    }
}
