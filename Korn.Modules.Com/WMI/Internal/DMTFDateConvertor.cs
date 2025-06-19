using System;
using System.Runtime.CompilerServices;

static unsafe class DMTFDateConvertor
{
    static readonly ushort[] DaysArray = [0, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365];
    static readonly ushort[] LeapDaysArray = [0, 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366];

    static readonly ushort* Days;
    static readonly ushort* LeapDays;

    static DMTFDateConvertor()
    {
        fixed (ushort* pointer = DaysArray, leapPointer = LeapDaysArray)
        {
            Days = pointer;
            LeapDays = leapPointer;
        }
    }

    public static DateTime ConvertToDateTime(ReadOnlySpan<char> bstr) => new((long)ConvertToTicks(bstr));

    public static DateTime ConvertToDateTime(char* bstr) => new((long)ConvertToTicks(bstr));

    public static ulong ConvertToTicks(ReadOnlySpan<char> bstr)
    {
        fixed (char* chars = bstr)
            return ConvertToTicks(chars);
    }

    [SkipLocalsInit]
    public static ulong ConvertToTicks(char* bstr/*wchar[25]*/)
    {
        ulong a, b, c, d;

        a = *(ulong*)&bstr[4];
        d = a & 15;
        d = (((d << 2) + d) << 1) + (a >> 16 & 15);

        a = *(ulong*)bstr;
        b = a & 15;
        b = (((b << 2) + b) << 1) + (a >> 16 & 15);
        b = (((b << 2) + b) << 1) + (a >> 32 & 15);
        b = (((b << 2) + b) << 1) + (a >> 48 & 15);
        a = b - 1;
        c = a * 1461 / 4 - a / 100 + a / 400 + ((b & 3) != 0 ? Days[d] : (b & 15) == 0 ? LeapDays[d] : b % 25 == 0 ? Days[d] : LeapDays[d]);

        a = *(ulong*)&bstr[19];
        d = a & 15;
        d = (((d << 2) + d) << 1) + (a >> 16 & 15);
        d = ((d << 2) + d) << 1;

        a = *(ulong*)&bstr[15];
        b = a & 15;
        b = (((b << 2) + b) << 1) + (a >> 16 & 15);
        b = (((b << 2) + b) << 1) + (a >> 32 & 15);
        b = (((b << 2) + b) << 1) + (a >> 48 & 15);
        b = ((b << 4) + (b << 3) + b) << 2;
        d += ((b << 2) + b) << 1;

        a = *(ulong*)&bstr[10];
        b = a & 15;
        b = (((b << 2) + b) << 1) + (a >> 16 & 15);
        d += (((b << 6) - (b << 2) + (a >> 48 & 15) + (a >> 32 & 15) * 10) << 7) * 0x1312D;

        a = *(ulong*)&bstr[6];
        b = a & 15;
        d += (((c + (a >> 16 & 15) + (((b << 2) + b) << 1) - 1) << 6) * 0x324A9A7 + (((a >> 48 & 15) + (a >> 32 & 15) * 10) << 3) * 0x10C388D) << 8;

        return d;
    }
}
