using System.Collections.Generic;

namespace BimFun
{
    internal class FileComparer : IComparer<string>
    {
        [System.Runtime.InteropServices.DllImport("Shlwapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string psz1, string psz2);
        public int Compare(string psz1, string psz2)
        {
            return StrCmpLogicalW(psz1, psz2);
        }
    }
}