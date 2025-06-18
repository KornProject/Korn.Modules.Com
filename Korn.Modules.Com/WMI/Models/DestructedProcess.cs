#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
namespace Korn.Com.Wmi;
public unsafe class DestructedProcess
{
    public int ID;       // UInt32 ProcessId 
    public int ParentID; // UInt32 ParentProcessId
    public string Name;  // BSTR   Name
}