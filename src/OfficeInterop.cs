// UnLocked - Office 自動化輔助類別（C# 5 語法，供 Windows PowerShell 5.1 Add-Type 編譯）
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace UnLocked
{
    /// <summary>
    /// 逾時看門狗：若單一檔案處理超過指定秒數（例如 Office 跳出看不見的對話框而卡住），
    /// 便強制結束由本工具啟動的 Office 處理程序，讓批次作業可以繼續處理下一個檔案。
    /// 以 System.Threading.Timer 實作，不依賴 PowerShell 執行空間。
    /// </summary>
    public sealed class Watchdog : IDisposable
    {
        readonly object gate = new object();
        readonly Timer timer;
        int pid;
        DateTime deadline;
        bool armed;
        volatile bool fired;

        public Watchdog()
        {
            timer = new Timer(Tick, null, 1000, 1000);
        }

        public bool Fired { get { return fired; } }

        public void Arm(int processId, int seconds)
        {
            lock (gate)
            {
                pid = processId;
                deadline = DateTime.UtcNow.AddSeconds(seconds);
                armed = processId > 0 && seconds > 0;
                fired = false;
            }
        }

        public void Disarm()
        {
            lock (gate) { armed = false; }
        }

        void Tick(object state)
        {
            int target;
            lock (gate)
            {
                if (!armed || DateTime.UtcNow < deadline) return;
                armed = false;
                fired = true;
                target = pid;
            }
            try { Process.GetProcessById(target).Kill(); }
            catch (Exception) { }
        }

        public void Dispose()
        {
            timer.Dispose();
        }
    }

    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool SetProcessDPIAware();

        /// <summary>隱藏啟動用的主控台視窗（GUI 模式）。</summary>
        public static void HideConsole()
        {
            IntPtr h = GetConsoleWindow();
            if (h != IntPtr.Zero) ShowWindow(h, 0);
        }

        public static void EnableDpiAwareness()
        {
            try { SetProcessDPIAware(); } catch (Exception) { }
        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>把 Office 視窗帶到最前面，讓使用者看得到登入／權限提示。</summary>
        public static void BringToFront(long hwnd)
        {
            if (hwnd == 0) return;
            IntPtr h = new IntPtr(hwnd);
            ShowWindow(h, 9);   // SW_RESTORE
            SetForegroundWindow(h);
        }

        public static int GetPidFromHwnd(long hwnd)
        {
            if (hwnd == 0) return 0;
            uint id;
            GetWindowThreadProcessId(new IntPtr(hwnd), out id);
            return (int)id;
        }
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    /// <summary>
    /// OLE 訊息篩選器：Office 忙碌時（RPC_E_CALL_REJECTED）自動重試，避免「被呼叫端拒絕呼叫」錯誤。
    /// 只能在 STA 執行緒上註冊。
    /// </summary>
    public sealed class MessageFilter : IOleMessageFilter
    {
        [DllImport("ole32.dll")]
        static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        public static bool Register()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) return false;
            IOleMessageFilter old;
            return CoRegisterMessageFilter(new MessageFilter(), out old) == 0;
        }

        public static void Revoke()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) return;
            IOleMessageFilter old;
            CoRegisterMessageFilter(null, out old);
        }

        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0; // SERVERCALL_ISHANDLED
        }

        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // SERVERCALL_RETRYLATER = 2：每 250 毫秒重試一次，最多 3 分鐘
            if (dwRejectType == 2 && dwTickCount < 180000) return 250;
            return -1;
        }

        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2; // PENDINGMSG_WAITDEFPROCESS
        }
    }
}
