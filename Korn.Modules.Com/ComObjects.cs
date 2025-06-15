using System;
using Korn.Utils;
using Korn.Com.Internal;
using System.Threading;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
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
        WbemObjectSink.Interface implementation)
    {
        fixed (char* queryLanguagePointer = queryLanguage)
            return ((delegate* unmanaged<WbemServices, char*, char*, uint, WbemContext, WbemObjectSink, int>)self[23])(
                self.As<WbemServices>(),
                queryLanguagePointer,
                query,
                flags,
                context,
                implementation.As<WbemObjectSink>()
            );
    }

    public static int CancelAsyncCall( this WbemServices.Interface self, WbemObjectSink.Interface implementation) 
        => ((delegate* unmanaged<WbemServices, WbemObjectSink, int>)self[4])(self.As<WbemServices>(), implementation.As<WbemObjectSink>());

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

    public static WbemObjectSinkImplementation* GetImplementation(this WbemObjectSink.Interface self)
        => ((delegate* unmanaged<WbemObjectSink, WbemObjectSinkImplementation*>)self[5])(self.As<WbemObjectSink>());
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
public unsafe struct WbemClassObject : WbemClassObject.Interface
{
    public static IID IID { get; private set; } = "dc12a681737f11cf884d00aa004b2e24"u8;
    public interface Interface : IUnknown.Interface;
}

public unsafe struct WbemObjectSinkImplementation
{
    const uint WBEM_S_NO_ERROR = 0x00000000;
    const uint E_NOINTERFACE = 0x80004002;

    void** vtable;
    uint refs;

    void Delete() => Marshal.FreeCoTaskMem((nint)vtable);

    [UnmanagedCallersOnly]
    static uint QueryInterface(WbemObjectSinkImplementation* self, IID* iid, void** ppv)
    {
        if (*iid == IUnknown.IID || *iid == WbemObjectSink.IID)
        {
            *ppv = self;
            return WBEM_S_NO_ERROR;
        }

        return E_NOINTERFACE;
    }

    [UnmanagedCallersOnly]
    static uint AddRef(WbemObjectSinkImplementation* self) => Interlocked.Increment(ref self->refs);

    [UnmanagedCallersOnly]
    static uint Release(WbemObjectSinkImplementation* self)
    {
        var refs = Interlocked.Decrement(ref self->refs);

        if (refs == 0)
            self->Delete();

        return refs;
    }

    [UnmanagedCallersOnly]
    static uint SetStatus(WbemObjectSinkImplementation* self, int flags, char* param, WbemClassObject array) => WBEM_S_NO_ERROR;

    [UnmanagedCallersOnly]
    static WbemObjectSinkImplementation* GetImplementation(WbemObjectSinkImplementation* self) => self;

    public void Initialize(void* indicateFunction)
    {
        refs = 0;

        var vtable = this.vtable = (void**)Marshal.AllocCoTaskMem(sizeof(void*) * 6);
        *vtable++ = (delegate* unmanaged<WbemObjectSinkImplementation*, IID*, void**, uint>)&QueryInterface;
        *vtable++ = (delegate* unmanaged<WbemObjectSinkImplementation*, uint>)&AddRef;
        *vtable++ = (delegate* unmanaged<WbemObjectSinkImplementation*, uint>)&Release;
        *vtable++ = indicateFunction;
        *vtable++ = (delegate* unmanaged<WbemObjectSinkImplementation*, int, char*, WbemClassObject, uint>)&SetStatus;
        *vtable++ = (delegate* unmanaged<WbemObjectSinkImplementation*, WbemObjectSinkImplementation*>)&GetImplementation;
    }

    public WbemObjectSink GetSink()
    {
        var self = Unsafe.AsPointer(ref this);
        return *(WbemObjectSink*)&self;
    }
}