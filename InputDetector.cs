using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Windows.Threading;

namespace OSK
{
    /// <summary>
    /// 全域輸入焦點偵測器 (精準判定與自動隱藏修正版)
    /// 修正：檔案總管完成輸入後不自動隱藏、FB/IG 圖片誤觸發問題。
    /// </summary>
    public class InputDetector
    {
        private readonly Window _keyboardWindow;
        private AutomationFocusChangedEventHandler? _focusHandler;
        private DispatcherTimer _caretCheckTimer;
        private bool _isCurrentInputActive = false;

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

        private const int GUI_CARETBLINKING = 0x00000001;
        #endregion

        public InputDetector(Window keyboardWindow)
        {
            _keyboardWindow = keyboardWindow;
            
            _caretCheckTimer = new DispatcherTimer();
            _caretCheckTimer.Interval = TimeSpan.FromMilliseconds(400); // 略微調快偵測頻率
            _caretCheckTimer.Tick += CaretCheckTimer_Tick;
        }

        public void Start()
        {
            try
            {
                _focusHandler = new AutomationFocusChangedEventHandler(OnFocusChanged);
                Automation.AddAutomationFocusChangedEventHandler(_focusHandler);
                _caretCheckTimer.Start();
                Debug.WriteLine("[OSK] 輸入偵測啟動...");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OSK] 啟動失敗: {ex.Message}");
            }
        }

        public void Stop()
        {
            if (_focusHandler != null)
            {
                Automation.RemoveAutomationFocusChangedEventHandler(_focusHandler);
                _focusHandler = null;
            }
            _caretCheckTimer.Stop();
        }

        /// <summary>
        /// 定時檢查系統是否有閃爍游標，這對於解決檔案總管完成輸入後隱藏非常重要。
        /// </summary>
        private void CaretCheckTimer_Tick(object? sender, EventArgs e)
        {
            GUITHREADINFO gui = new GUITHREADINFO();
            gui.cbSize = Marshal.SizeOf(gui);

            if (GetGUIThreadInfo(0, ref gui))
            {
                // 檢查是否有實體 Caret 或正在閃爍
                bool hasCaret = gui.hwndCaret != IntPtr.Zero || (gui.flags & GUI_CARETBLINKING) != 0;

                if (hasCaret)
                {
                    uint pid;
                    GetWindowThreadProcessId(GetForegroundWindow(), out pid);
                    if (pid == Process.GetCurrentProcess().Id) return;

                    _isCurrentInputActive = true;
                    ShowKeyboard();
                }
                else
                {
                    // 如果 UIA 焦點目前也不在輸入區，則根據 Caret 消失來隱藏鍵盤
                    if (!_isCurrentInputActive)
                    {
                        HideKeyboard();
                    }
                }
            }
        }

        private void OnFocusChanged(object sender, AutomationFocusChangedEventArgs e)
        {
            try
            {
                if (sender is AutomationElement element)
                {
                    var current = element.Current;
                    var controlType = current.ControlType;
                    var className = current.ClassName ?? "";
                    var name = current.Name ?? "";

                    // 1. 系統組件過濾 (排除開始功能表、工作列)
                    if (className.Contains("Shell_") || className.Contains("Tray") || name == "開始" || name == "Search")
                    {
                        _isCurrentInputActive = false;
                        HideKeyboard();
                        return;
                    }

                    // 2. 狀態檢查
                    if (!current.IsEnabled || current.IsOffscreen)
                    {
                        _isCurrentInputActive = false;
                        HideKeyboard();
                        return;
                    }

                    bool isInputArea = false;

                    // 3. 判定邏輯
                    if (current.IsPassword || controlType == ControlType.Edit || controlType == ControlType.ComboBox)
                    {
                        isInputArea = true;
                    }
                    else if (controlType == ControlType.Document)
                    {
                        // 處理瀏覽器/社群網站，檢查是否唯讀
                        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object vp))
                        {
                            var valPattern = (ValuePattern)vp;
                            if (!valPattern.Current.IsReadOnly) isInputArea = true;
                        }
                        else if (element.TryGetCurrentPattern(TextPattern.Pattern, out _) && current.IsKeyboardFocusable)
                        {
                            if (!name.Contains("圖片") && !name.Contains("Photo"))
                            {
                                isInputArea = true;
                            }
                        }
                    }
                    else
                    {
                        // 針對 LINE 或 wxMedit 等非標準 UI
                        if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object p))
                        {
                            var vp = (ValuePattern)p;
                            if (!vp.Current.IsReadOnly && current.IsKeyboardFocusable)
                            {
                                isInputArea = true;
                            }
                        }
                    }

                    _isCurrentInputActive = isInputArea;

                    if (isInputArea)
                    {
                        uint pid;
                        GetWindowThreadProcessId(GetForegroundWindow(), out pid);
                        if (pid == Process.GetCurrentProcess().Id) return;

                        ShowKeyboard();
                    }
                    else
                    {
                        // 當焦點移動到非輸入區(例如檔案總管的空白處)，執行隱藏
                        HideKeyboard();
                    }
                }
            }
            catch (ElementNotAvailableException) 
            {
                // 當元素消失時（如檔案總管完成重新命名），視為需要隱藏
                _isCurrentInputActive = false;
                HideKeyboard();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OSK] UIA 錯誤: {ex.Message}");
            }
        }

        private void ShowKeyboard()
        {
            _keyboardWindow.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_keyboardWindow.Visibility != Visibility.Visible)
                {
                    _keyboardWindow.Show();
                }
                if (_keyboardWindow.WindowState == WindowState.Minimized)
                    _keyboardWindow.WindowState = WindowState.Normal;
                
                _keyboardWindow.Topmost = true;
            }), DispatcherPriority.Input);
        }

        private void HideKeyboard()
        {
            _keyboardWindow.Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_keyboardWindow.Visibility == Visibility.Visible)
                {
                    _keyboardWindow.Hide();
                }
            }), DispatcherPriority.Input);
        }
    }
}