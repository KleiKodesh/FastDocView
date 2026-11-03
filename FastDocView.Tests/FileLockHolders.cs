using System.Diagnostics;
using System.Runtime.InteropServices;

namespace FastDocView.Tests;

/// <summary>
/// Which processes hold a file open, from the Windows Restart Manager (RmGetList) — the same query installers use
/// to name the program that locks a file. Used to tell our own handle from another process's (an antivirus
/// scanning a freshly written file) when a sharing check fails.
/// </summary>
internal static class FileLockHolders
{
    [StructLayout(LayoutKind.Sequential)]
    private struct UniqueProcess { public int ProcessId; public System.Runtime.InteropServices.ComTypes.FILETIME StartTime; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessInfo
    {
        public UniqueProcess Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string AppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string ServiceShortName;
        public int ApplicationType;
        public uint AppStatus;
        public uint TsSessionId;
        [MarshalAs(UnmanagedType.Bool)] public bool Restartable;
    }

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)] private static extern int RmStartSession(out uint handle, int flags, string key);
    [DllImport("rstrtmgr.dll")] private static extern int RmEndSession(uint handle);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(uint handle, uint fileCount, string[] files, uint applicationCount, UniqueProcess[]? applications, uint serviceCount, string[]? services);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmGetList(uint handle, out uint needed, ref uint count, [In, Out] ProcessInfo[]? processes, ref uint rebootReasons);

    /// <summary>(process id, name) of every process holding one of the files open; empty when none (or the query failed: see the error text).</summary>
    public static List<(int processId, string name)> Of(params string[] paths)
    {
        var holders = new List<(int, string)>();
        if (RmStartSession(out uint session, 0, Guid.NewGuid().ToString("N")) != 0) return holders;
        try
        {
            if (RmRegisterResources(session, (uint)paths.Length, paths, 0, null, 0, null) != 0) return holders;
            uint count = 0, reasons = 0;
            int result = RmGetList(session, out uint needed, ref count, null, ref reasons);
            if (result != 234 || needed == 0) return holders;          // ERROR_MORE_DATA: someone holds it
            var processes = new ProcessInfo[needed];
            count = needed;
            if (RmGetList(session, out needed, ref count, processes, ref reasons) != 0) return holders;
            for (int i = 0; i < count; i++)
            {
                string name = processes[i].AppName;
                try { name = Process.GetProcessById(processes[i].Process.ProcessId).ProcessName + " (" + name + ")"; } catch (ArgumentException) { }
                holders.Add((processes[i].Process.ProcessId, name));
            }
            return holders;
        }
        finally { RmEndSession(session); }
    }
}
