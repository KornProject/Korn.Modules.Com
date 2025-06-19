#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider adding the 'required' modifier or declaring as nullable.
namespace Korn.Modules.Com.Wmi;
public unsafe class CreatedProcess
{
    public int ID;                // UInt32 ProcessId 
    public int ParentID;          // UInt32 ParentProcessId
    public ulong CreationTime;    // UInt64 CreationDate 
    public string Name;           // BSTR   Name
    public string CommandLine;    // BSTR   CommandLine
    public string ExecutablePath; // BSTR   ExecutablePath
}