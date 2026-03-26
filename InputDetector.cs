using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Text;
using System.IO;

namespace OSK
{
    /// <summary>
    /// 全域輸入焦點偵測器 (診斷強化版)
    /// 加入實體日誌以追蹤管理員視窗判定失敗的原因。
    /// </summary>
    public class InputDetector
    {
        private readonly Window _keyboardWindow;
        private AutomationFocusChangedEventHandler? _focusHandler;
        private DispatcherTimer _backupTimer;
        private bool _isCurrentInputActive = false;
        private readonly int _currentProcessId;
        
        private IntPtr _hWinEventHook;
        private WinEventProc _winEventDelegate;
        private string _logPath;

        #region Win32 API
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventProc lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private delegate void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        private const int GUI_CARETBLINKING = 0x00000001;
        #endregion

        public InputDetector(Window keyboardWindow)
        {
            _keyboardWindow = keyboardWindow;
            using (var proc = Process.GetCurrentProcess())
            {
                _currentProcessId = proc.Id;
            }
            
            _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "osk_debug.log");
            
            _backupTimer = new DispatcherTimer();
            _backupTimer.Interval = TimeSpan.FromMilliseconds(500);
            _backupTimer.Tick += (s, e) => CheckForegroundAndCaret("Timer");

            _winEventDelegate = new WinEventProc(WinEventCallback);
        }

        private void Log(string message)
        {
            try
            {
                File.AppendAllText(_logPath, $"{DateTime.Now:HH:mm:ss.fff} {message}\n");
            }
            catch { }
        }

        public void Start()
        {
            if (_focusHandler != null) return; // 避免重複啟動導致掛鉤重疊

            Log("--- Detection Started ---");
            try
            {
                _focusHandler = new AutomationFocusChangedEventHandler(OnFocusChanged);
                Automation.AddAutomationFocusChangedEventHandler(_focusHandler);

                _hWinEventHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _winEventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
                
                _backupTimer.Start();
            }
            catch (Exception ex)
            {
                Log($"Start Failed: {ex.Message}");
                _focusHandler = null;
            }
        }

        public void Stop(bool isShuttingDown = false)
        {
            // [關鍵優化] 關閉程式時，UIA 的 Remove 操作極易因為 COM 狀態而導致執行緒掛死
            // 如果是正在關閉進程，我們跳過 UIA 註銷，讓作業系統在進程結束時統一回收
            if (!isShuttingDown)
            {
                var handler = _focusHandler;
                if (handler != null)
                {
                    _focusHandler = null; 
                    try { Automation.RemoveAutomationFocusChangedEventHandler(handler); } catch { }
                }
            }
            else
            {
                _focusHandler = null; // 標記為空即可
            }

            if (_hWinEventHook != IntPtr.Zero)
            {
                UnhookWinEvent(_hWinEventHook);
                _hWinEventHook = IntPtr.Zero;
            }
            _backupTimer.Stop();
            Log("--- Detection Stopped ---");
        }

        private void WinEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            CheckForegroundAndCaret("WinEvent");
        }

        private void CheckForegroundAndCaret(string source)
        {
            if (_keyboardWindow is MainWindow mw && mw.IsBusySuppressing) return;

            IntPtr hFore = GetForegroundWindow();
            if (hFore == IntPtr.Zero) return;

            uint pid;
            uint tid = GetWindowThreadProcessId(hFore, out pid);
            if (pid == (uint)_currentProcessId) return;

            var className = new StringBuilder(256);
            GetClassName(hFore, className, className.Capacity);
            string cls = className.ToString();
            string clsLower = cls.ToLower();

            // 如果是管理員視窗類別，強制狀態為 Active (包含 Windows Terminal 的 CASCADIA_HOSTING_WINDOW_CLASS)
            if (clsLower.Contains("terminal") || clsLower.Contains("console") || 
                clsLower.Contains("cmd") || clsLower.Contains("powershell") || clsLower.Contains("cascadia"))
            {
                Log($"[{source}] Admin Window Detected: {cls}");
                UpdateState(true);
                return;
            }

            // 標準 Win32 Caret 偵測
            GUITHREADINFO gui = new GUITHREADINFO();
            gui.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));
            if (GetGUIThreadInfo(tid, ref gui))
            {
                if (gui.hwndCaret != IntPtr.Zero || (gui.flags & GUI_CARETBLINKING) != 0)
                {
                    Log($"[{source}] Caret Detected in: {cls}");
                    UpdateState(true);
                }
                else
                {
                    if (clsLower.Contains("progman") || clsLower.Contains("workerw"))
                    {
                        Log($"[{source}] Desktop Detected, Hiding.");
                        UpdateState(false);
                    }
                }
            }
            else
            {
                // 如果 GetGUIThreadInfo 失敗 (通常是提權視窗)，但類別沒被上面軌道抓到
                Log($"[{source}] GUI Info Failed for: {cls} (PID: {pid})");
            }
        }

        private void OnFocusChanged(object sender, AutomationFocusChangedEventArgs e)
        {
            if (_keyboardWindow is MainWindow mw && mw.IsBusySuppressing) return;

            try
            {
                if (sender is AutomationElement element)
                {
                    var current = element.Current;
                    var controlType = current.ControlType;
                    var className = current.ClassName ?? "";
                    
                    if (className.Contains("Shell_") || className.Contains("Tray"))
                    {
                        UpdateState(false);
                        return;
                    }

                    bool isInputArea = false;
                    if (current.IsPassword || controlType == ControlType.Edit || controlType == ControlType.ComboBox)
                    {
                        isInputArea = true;
                    }
                    else if (controlType == ControlType.Document)
                    {
                        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object vp))
                        {
                            if (!((ValuePattern)vp).Current.IsReadOnly) isInputArea = true;
                        }
                    }

                    if (isInputArea) Log($"[UIA] Input Area: {controlType.ProgrammaticName} in {className}");
                    UpdateState(isInputArea);
                }
            }
            catch (Exception)
            {
                // UIA 跨進程失敗很正常，不記日誌以免刷屏
            }
        }

        private void UpdateState(bool active)
        {
            _isCurrentInputActive = active;
            _keyboardWindow.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (active)
                {
                    if (_keyboardWindow.Visibility != Visibility.Visible)
                    {
                        _keyboardWindow.Show();
                        Log("-> Keyboard Show");
                    }
                    if (_keyboardWindow.WindowState == WindowState.Minimized)
                        _keyboardWindow.WindowState = WindowState.Normal;
                    _keyboardWindow.Topmost = true;
                }
                else
                {
                    if (_keyboardWindow.Visibility == Visibility.Visible)
                    {
                        _keyboardWindow.Hide();
                        Log("-> Keyboard Hide");
                    }
                }
            }), DispatcherPriority.Input);
        }
    }
}