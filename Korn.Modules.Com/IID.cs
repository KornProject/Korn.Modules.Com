using System;
using System.Buffers.Text;

#pragma warning disable CS0660 // Type defines operator == or operator != but does not override Object.Equals(object o)
#pragma warning disable CS0661 // Type defines operator == or operator != but does not override Object.GetHashCode()
namespace Korn.Modules.Com;
public unsafe struct IID
{
    public IID(Guid guid) => GUID = guid;
    public IID(ReadOnlySpan<byte> u8string)
    {
        if (!Utf8Parser.TryParse(u8string, out Guid guid, out _, 'N'))
            throw new Exception("Unable to parse guid in IID");

        GUID = guid;
    }

    public Guid GUID;

    public static implicit operator IID(ReadOnlySpan<byte> u8string) => new(u8string);
    public static bool operator ==(IID left, IID right) => left.GUID == right.GUID;
    public static bool operator !=(IID left, IID right) => left.GUID != right.GUID;
}