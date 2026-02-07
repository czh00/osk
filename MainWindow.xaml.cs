using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.ComponentModel;
using System.Windows.Controls;
using System.Drawing;
using System.Linq;
using System.Windows.Threading;
using System.Threading;
using System.Windows.Media;
using System.Reflection; // 新增 Reflection 引用

namespace OSK
{
    /// <summary>
    /// 主視窗：一個可停靠、非啟動激活的虛擬鍵盤 (WPF)。
    /// 功能總覽：
    /// - 顯示虛擬按鍵、處理按鍵點擊並注入虛擬按鍵事件 (keybd_event)。
    /// - 使用全域鍵盤掛鉤 (WH_KEYBOARD_LL) 追蹤實體鍵盤狀態以同步按鍵按下/放開視覺效果。
    /// - 偵測前景視窗的 IME 狀態 (使用 IMM32 API) 以決定是否顯示注音或英文字母。
    /// - 支援「Mode」鍵以切換小狼毫/注音狀態（會傳送 Shift 以觸發系統 IME 切換），並支援本地預覽模式避免變更系統 IME。
    /// - 支援「臨時英文模式」：在注音狀態按 Shift 啟動，第一個字母以大寫送出，後續字母以小寫送出，Enter 結束該臨時模式。
    /// - 支援虛擬 modifier 鍵 (Shift/Ctrl/Alt/Win) 與 Fn 模式（Fn 可改變鍵盤顯示與發送行為）。
    /// 注意事項：
    /// - 本類大量使用 Win32 API 與 IME API，僅在 Windows 平台有效。
    /// - 由於使用低階鍵盤掛鉤與注入按鍵，需注意不要與系統或其他輸入攔截器產生衝突。
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        #region Win32 / IMM API 宣告
        // 以下為與 Win32/IMM 互動所需的 P/Invoke 宣告與結構。
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        /// <summary>
        /// 用於注入鍵盤事件 (down/up)。此 API 在現代 Windows 上仍可用，但會受安全性/系統設定影響。
        /// </summary>
        [DllImport("user32.dll")] private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        [DllImport("user32.dll")] private static extern uint RegisterWindowMessage(string lpString);
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // 用於替代 SAS 選單的系統操作
        [DllImport("user32.dll", SetLastError = true)] private static extern bool LockWorkStation();
        [DllImport("user32.dll", SetLastError = true)] private static extern bool ExitWindowsEx(uint uFlags, uint dwReason);

        /// <summary>
        /// 取得執行緒 GUI 狀態，能讓我們找到實際擁有輸入焦點的控制項 hwnd。
        /// </summary>
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

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
            public System.Drawing.Rectangle rcCaret;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        // IMM32 API：用來檢查/控制前景輸入法 (IME) 轉換狀態，判定是否為中文模式 (NATIVE)。
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmGetConversionStatus(IntPtr hIMC, out uint lpdwConversion, out uint lpdwSentence);
        [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmSetConversionStatus(IntPtr hIMC, uint dwConversion, uint dwSentence);

        // 低階鍵盤掛鉤相關
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string lpModuleName);
        #endregion

        #region 常數與委派
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const uint IME_CMODE_NATIVE = 0x0001; // IMM: native bit 表示中文輸入
        private const byte MODE_KEY_CODE = 0xFF; // 自定義的 Mode 虛擬鍵值 (不對應系統 VK)
        private const byte FN_KEY_CODE = 0xFE;   // 自定義的 Fn 虛擬鍵值
        #endregion

        #region 欄位（狀態與 UI 綁定）
        private LowLevelKeyboardProc? _proc;
        private IntPtr _hookID = IntPtr.Zero;
        private uint _msgShowOsk;
        private System.Windows.Forms.NotifyIcon _notifyIcon;

        // 鍵盤資料模型（每列按鍵）
        public ObservableCollection<ObservableCollection<KeyModel>> KeyRows { get; set; } = new();
        public ICommand KeyCommand { get; set; }

        // 目前系統 / 視覺狀態
        private bool _isZhuyinMode = false;        // 代表目前系統 IME 是否為中文注音模式（同步自 IMM）
        private bool _isShiftActive = false;
        private bool _isCapsLockActive = false;
        private bool _isCtrlActive = false;
        private bool _isAltActive = false;
        private bool _isWinActive = false;

        // 本地虛擬 modifier 與功能鍵切換（點擊虛擬鍵時保留狀態）
        private bool _virtualShiftToggle = false;
        private bool _virtualCtrlToggle = false;
        private bool _virtualAltToggle = false;
        private bool _virtualWinToggle = false;
        private bool _virtualFnToggle = false;

        // 本地預覽開關：當 true 時，介面顯示會反轉 IME 狀態，但不會實際變更系統 IME（方便校正視覺）
        private bool _localPreviewToggle = false;

        // 臨時英文模式支援：
        // - 由注音狀態中按下 Shift 啟動；第一個字母會以 Shift+letter 注入（大寫），之後字母以小寫送出，按 Enter 結束。
        private bool _temporaryEnglishMode = false;
        private bool _temporaryEnglishFirstUpperSent = false;
        private bool _temporaryShiftInjected = false;

        // 用於忽略程序注入的 Shift 事件（避免 HookCallback 誤判）
        private int _suppressInjectedShiftCount = 0;

        // 當使用 Mode 鍵切換 IME 時，短時間內忽略系統 IME 的回報（避免剛注入的 Shift 被回讀覆蓋）
        private DateTime _ignoreImeSyncUntil = DateTime.MinValue;

        // 指示器顯示字串與顏色（綁定至 UI）
        private string _modeIndicator = "En";
        public string ModeIndicator { get { return _modeIndicator; } set { _modeIndicator = value; OnPropertyChanged("ModeIndicator"); } }
        private string _indicatorColor = "White";
        public string IndicatorColor { get { return _indicatorColor; } set { _indicatorColor = value; OnPropertyChanged("IndicatorColor"); } }

        // Fn 模式顯示與實際發送對照表
        // 新增：當 Fn 模式啟用時，將倒退鍵 (Backspace, VK=0x08) 顯示為 "Del" 並發送 Delete (VK=0x2E)
        private static readonly Dictionary<byte, string?> FnDisplayMap = new()
        {
            [0xC0] = "⎋",
            [0x31] = "F1",
            [0x32] = "F2",
            [0x33] = "F3",
            [0x34] = "F4",
            [0x35] = "F5",
            [0x36] = "F6",
            [0x37] = "F7",
            [0x38] = "F8",
            [0x39] = "F9",
            [0x30] = "F10",
            [0xBD] = "F11",
            [0xBB] = "F12",
            [0x26] = "⎗",
            [0x28] = "⎘",
            [0x25] = "⌂",
            [0x27] = "⤓",
            [0x09] = "⌃⌥⌦", // Tab (in Fn mode) -> show Ctrl+Alt+Del symbol
            [0x08] = "⌦"    // Backspace -> Del 顯示於 Fn 模式
        };

        private static readonly Dictionary<byte, byte> FnSendMap = new()
        {
            [0xC0] = 0x1B,
            [0x31] = 0x70,
            [0x32] = 0x71,
            [0x33] = 0x72,
            [0x34] = 0x73,
            [0x35] = 0x74,
            [0x36] = 0x75,
            [0x37] = 0x76,
            [0x38] = 0x77,
            [0x39] = 0x78,
            [0x30] = 0x79,
            [0xBD] = 0x7A,
            [0xBB] = 0x7B,
            [0x26] = 0x21,
            [0x28] = 0x22,
            [0x25] = 0x24,
            [0x27] = 0x23,
            [0x09] = 0x2E,   // Tab (in Fn mode) -> send Delete (together with Ctrl+Alt semantics)
            [0x08] = 0x2E    // Backspace (0x08) -> Delete (0x2E) 在 Fn 模式下發送
        };

        private bool _isResizing = false;
        #endregion

        /// <summary>
        /// 建構子：初始化 UI 綁定、鍵盤模型、系統通知圖示、全域鍵盤掛鉤與週期性狀態同步計時器。
        /// - KeyCommand 綁定到虛擬鍵點擊處理器 OnKeyClick。
        /// - 使用 RegisterWindowMessage 取得自訂訊息以控制顯示/隱藏。
        /// - 建立 NotifyIcon（系統托盤）並註冊簡單的上下文選單。
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;

            // 設定為手動啟動位置，並在 Loaded 時把視窗置中並貼齊下緣
            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Loaded += MainWindow_Loaded;

            // 若 XAML 有一個名為 OpacitySlider 的 Slider，啟用點擊定位行為 (IsMoveToPointEnabled)，
            // 並提供回退的 PreviewMouseLeftButtonDown 處理器確保點擊軌道能精準設值。
            var opacitySlider = this.FindName("OpacitySlider") as System.Windows.Controls.Slider;
            if (opacitySlider != null)
            {
                try
                {
                    opacitySlider.IsMoveToPointEnabled = true;
                }
                catch { }
                opacitySlider.PreviewMouseLeftButtonDown += OpacitySlider_PreviewMouseLeftButtonDown;
            }

            KeyCommand = new RelayCommand<KeyModel>(OnKeyClick);
            SetupKeyboard();

            // KeyBoardItemsControl 應在 XAML 中定義，其 ItemsSource 綁定在此。
            KeyBoardItemsControl.ItemsSource = KeyRows;

            _msgShowOsk = RegisterWindowMessage("WM_SHOW_OSK_V100");

            // --- 修正：載入系統托盤圖示 ---
            // 策略：優先從執行檔本身提取圖示 (ExtractAssociatedIcon)，這樣最能保證與工作列一致。
            System.Drawing.Icon? trayIcon = null;

            try
            {
                // 取得目前執行檔的路徑 (Process.GetCurrentProcess().MainModule.FileName 比 Assembly.Location 更可靠)
                string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                
                if (!string.IsNullOrEmpty(exePath) && System.IO.File.Exists(exePath))
                {
                    // 直接從 EXE 提取關聯圖示
                    trayIcon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"圖示提取失敗: {ex.Message}");
            }

            // Fallback: 如果從 EXE 提取失敗，嘗試載入 Resource (針對開發環境或特殊打包)
            if (trayIcon == null)
            {
                try
                {
                    var resourceUri = new Uri("pack://application:,,,/Assets/app.ico");
                    var streamInfo = System.Windows.Application.GetResourceStream(resourceUri);
                    if (streamInfo != null)
                    {
                        trayIcon = new System.Drawing.Icon(streamInfo.Stream);
                    }
                }
                catch { /* 忽略 */ }
            }

            // Last resort: 使用系統預設圖示
            if (trayIcon == null)
            {
                trayIcon = SystemIcons.Application;
            }

            // 建立系統托盤圖示與選單
            _notifyIcon = new System.Windows.Forms.NotifyIcon { Icon = trayIcon, Visible = true, Text = "OSK v1.0.0" };
            _notifyIcon.Click += (s, e) => ToggleVisibility();
            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add("結束", null, (s, e) => System.Windows.Application.Current.Shutdown());
            _notifyIcon.ContextMenuStrip = menu;

            // 設置低階鍵盤掛鉤以同步實體鍵盤狀態（視覺反饋）
            _proc = HookCallback;
            _hookID = SetHook(_proc);

            // 週期性同步：更新 Caps/Shift/Ctrl/Alt/Win 與 IME 狀態並刷新鍵盤顯示
            DispatcherTimer timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
            timer.Tick += (s, e) => SyncStates();
            timer.Start();
        }

        /// <summary>
        /// Loaded 時把視窗水平置中並貼齊該螢幕工作區下緣。
        /// 會使用游標所在的螢幕（多螢幕友善）。
        /// 重要：把像素座標轉成 WPF 裝置無關單位以避免 DPI 導致跑到畫面外。
        /// </summary>
        private void MainWindow_Loaded(object? sender, RoutedEventArgs e)
        {
            // 取用游標所在螢幕的工作區（像素）
            var cursorPos = System.Windows.Forms.Cursor.Position;
            var screen = System.Windows.Forms.Screen.FromPoint(cursorPos);
            var wa = screen.WorkingArea; // System.Drawing.Rectangle (像素)

            // 取得 device -> WPF 的轉換矩陣（如果可用）
            var source = PresentationSource.FromVisual(this);
            Matrix transform = Matrix.Identity;
            if (source?.CompositionTarget != null) transform = source.CompositionTarget.TransformFromDevice;

            // 把工作區像素座標轉為 WPF 單位
            var topLeft = transform.Transform(new System.Windows.Point(wa.Left, wa.Top));
            var bottomRight = transform.Transform(new System.Windows.Point(wa.Right, wa.Bottom));
            double waLeft = topLeft.X;
            double waTop = topLeft.Y;
            double waWidth = Math.Max(0, bottomRight.X - topLeft.X);
            double waHeight = Math.Max(0, bottomRight.Y - topLeft.Y);

            // 使用實際呈現大小（若尚未量測則 fallback 為設計寬高）
            double w = this.ActualWidth > 0 ? this.ActualWidth : (double.IsNaN(this.Width) ? 1000 : this.Width);
            double h = this.ActualHeight > 0 ? this.ActualHeight : (double.IsNaN(this.Height) ? 360 : this.Height);

            // 計算置中與貼齊下緣的位置（使用 WPF 單位）
            this.Left = waLeft + (waWidth - w) / 2.0;
            this.Top = waTop + (waHeight - h);

            // 邊界檢查，確保視窗至少有部分可見；如果超出則微調回到可視範圍
            if (this.Left < waLeft) this.Left = waLeft;
            if (this.Left + w > waLeft + waWidth) this.Left = waLeft + waWidth - w;
            if (this.Top < waTop) this.Top = waTop;
            if (this.Top + h > waTop + waHeight) this.Top = waTop + waHeight - h;
        }

        /// <summary>
        /// 建立低階鍵盤掛鉤 (WH_KEYBOARD_LL)。
        /// </summary>
        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule!.ModuleName!), 0);
            }
        }

        /// <summary>
        /// 低階鍵盤掛鉤回呼：用於同步 KeyRows 的按鍵按下/放開狀態、管理臨時英文模式及忽略程式注入所產生的事件。
        /// - 會忽略標記為注入或被 _suppressInjectedShiftCount 設定要忽略的 Shift 事件。
        /// - 按下實體 Shift 且目前為注音模式時，進入臨時英文模式（first-upper 的邏輯）。
        /// - 若注入 Shift（以送出第一個大寫字母），會在適當時機釋放注入的 Shift 並更新狀態。
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var kbd = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                int vk = (int)kbd.vkCode;
                uint flags = kbd.flags;
                bool injected = (flags & 0x10u) != 0; // LLKHF_INJECTED
                bool pressed = (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)0x0104);
                bool released = (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)0x0105);

                // 忽略程式注入或被 suppress 記號標示的 Shift 事件，以避免誤判
                if ((vk == 0x10 || vk == 0xA0 || vk == 0xA1) && (injected || _suppressInjectedShiftCount > 0))
                {
                    if (_suppressInjectedShiftCount > 0) _suppressInjectedShiftCount--;
                    return CallNextHookEx(_hookID, nCode, wParam, lParam);
                }

                if (pressed)
                {
                    // 若在注音模式下偵測到實體 Shift 按下，啟動臨時英文模式（只影響本工具的顯示與注入邏輯）
                    if ((vk == 0x10 || vk == 0xA0 || vk == 0xA1) && _isZhuyinMode)
                    {
                        _temporaryEnglishMode = true;
                        _temporaryEnglishFirstUpperSent = false;
                        UpdateDisplay();
                    }

                    // 臨時英文模式：若第一個按鍵是字母且使用者沒有按住實體 Shift，注入一個 Shift down 以送出大寫第一字
                    if (_temporaryEnglishMode)
                    {
                        bool isAlpha = vk >= 0x41 && vk <= 0x5A;
                        bool physicalShiftDown = (GetKeyState(0x10) & 0x8000) != 0;
                        if (isAlpha && !_temporaryEnglishFirstUpperSent && !physicalShiftDown)
                        {
                            _suppressInjectedShiftCount += 2; // 忽略注入的 down + up
                            keybd_event(0x10, 0, 0, 0); // 注入 Shift down
                            _temporaryShiftInjected = true;
                        }

                        // 若在臨時英文模式下按下 Enter，結束臨時模式
                        if (vk == 0x0D)
                        {
                            _temporaryEnglishMode = false;
                            _temporaryEnglishFirstUpperSent = false;
                            UpdateDisplay();
                        }
                    }
                }

                if (released)
                {
                    // 若之前注入了 Shift 且使用者放開的按鍵為英文字母，釋放注入的 Shift 並標記第一個大寫已送出
                    bool isAlpha = vk >= 0x41 && vk <= 0x5A;
                    if (_temporaryShiftInjected && isAlpha)
                    {
                        keybd_event(0x10, 0, 2, 0); // 注入 Shift up
                        _temporaryShiftInjected = false;
                        _temporaryEnglishFirstUpperSent = true;
                        UpdateDisplay();
                    }
                }

                // 同步視覺：將 KeyRows 中對應的按鍵標示為按下或放開
                if (pressed || released)
                {
                    foreach (var row in KeyRows)
                        foreach (var k in row)
                            if (k.VkCode == vk || ((vk == 0xA0 || vk == 0xA1) && k.VkCode == 0x10)) k.IsPressed = pressed;
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        /// <summary>
        /// 週期性狀態同步：
        /// - 同步 CapsLock / Shift / Ctrl / Alt / Win 真實鍵盤狀態並結合虛擬切換狀態。
        /// - 呼叫 DetectImeStatus() 來同步系統 IME 狀態。
        /// - 最後呼叫 UpdateDisplay() 來更新鍵盤視覺。
        /// </summary>
        private void SyncStates()
        {
            _isCapsLockActive = (GetKeyState(0x14) & 0x0001) != 0;
            _isShiftActive = ((GetKeyState(0x10) & 0x8000) != 0) || _virtualShiftToggle;
            _isCtrlActive = ((GetKeyState(0x11) & 0x8000) != 0) || _virtualCtrlToggle;
            _isAltActive = ((GetKeyState(0x12) & 0x8000) != 0) || _virtualAltToggle;
            _isWinActive = ((GetKeyState(0x5B) & 0x8000) != 0) || _virtualWinToggle;

            DetectImeStatus();
            UpdateDisplay();
        }

        /// <summary>
        /// 偵測前景視窗的 IME 轉換狀態（是否為中文/注音模式）。
        /// 流程：
        /// 1. 取得前景視窗與其執行緒 id。
        /// 2. 盡可能查找實際擁有焦點的控制項 hwnd（使用 GetGUIThreadInfo）。
        /// 3. 取得該 hwnd 的 IME window（ImmGetDefaultIMEWnd）並取得 IMC。
        /// 4. 使用 ImmGetConversionStatus() 判斷是否包含 IME_CMODE_NATIVE。
        /// 注意：當我們剛剛由 Mode 鍵切換 IME 時，會在短時間內忽略系統 IME 的回報 (_ignoreImeSyncUntil)。
        /// </summary>
        private void DetectImeStatus()
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero) return;

            uint threadId = GetWindowThreadProcessId(foregroundHwnd, out _);
            GUITHREADINFO guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = Marshal.SizeOf(guiInfo);

            IntPtr targetHwnd = foregroundHwnd;
            if (GetGUIThreadInfo(threadId, ref guiInfo))
            {
                if (guiInfo.hwndFocus != IntPtr.Zero) targetHwnd = guiInfo.hwndFocus;
            }

            IntPtr imeWnd = ImmGetDefaultIMEWnd(targetHwnd);
            IntPtr hIMC = ImmGetContext(imeWnd != IntPtr.Zero ? imeWnd : targetHwnd);

            if (hIMC != IntPtr.Zero)
            {
                if (DateTime.UtcNow < _ignoreImeSyncUntil)
                {
                    ImmReleaseContext(targetHwnd, hIMC);
                    return;
                }

                if (ImmGetConversionStatus(hIMC, out uint conv, out uint sentence))
                {
                    bool isChinese = (conv & IME_CMODE_NATIVE) != 0;
                    if (_isZhuyinMode != isChinese)
                    {
                        _isZhuyinMode = isChinese;
                        // 當系統 IME 真正變更時，清除本地預覽，確保視覺與系統一致
                        _localPreviewToggle = false;
                    }
                }
                ImmReleaseContext(targetHwnd, hIMC);
            }
        }

        /// <summary>
        /// 在視窗來源初始化時，設定視窗為不激活 (WS_EX_NOACTIVATE) 並註冊 WndProc 钩子以接收自訂訊息。
        /// </summary>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var helper = new WindowInteropHelper(this);
            SetWindowLong(helper.Handle, GWL_EXSTYLE, GetWindowLong(helper.Handle, GWL_EXSTYLE) | WS_EX_NOACTIVATE);
            HwndSource.FromHwnd(helper.Handle).AddHook(WndProc);
        }

        /// <summary>
        /// 視窗訊息處理：處理自訂的 _msgShowOsk 以切換顯示。
        /// </summary>
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == _msgShowOsk) { ToggleVisibility(); handled = true; }
            return IntPtr.Zero;
        }

        /// <summary>
        /// 切換視窗顯示/隱藏。
        /// </summary>
        private void ToggleVisibility()
        {
            if (this.Visibility == Visibility.Visible) this.Hide();
            else { this.Show(); }
        }

        /// <summary>
        /// 虛擬鍵點擊事件處理器：
        /// - Mode 鍵：若同時啟用 Fn，則只切換本地預覽；否則注入一對 Shift 以切換系統 IME（小狼毫行為），並短暫忽略 IME 同步。
        /// - Fn/Ctrl/Alt/Win 鍵：切換對應的虛擬開關，影響後續按鍵的發送或顯示。
        /// - Shift：在注音模式下以啟動臨時英文模式；否則切換虛擬 Shift。
        /// - 一般鍵：根據當前狀態（虛擬 modifiers、Fn map、臨時英文模式），組合並注入相應的按鍵事件。
        /// - 臨時英文模式：第一個英文字母以 Shift+letter 發出（若使用者沒按實體 Shift，會由程式注入），之後字母以單字母發出，按 Enter 結束模式。
        /// 特別處理：Ctrl+Alt+Fn+Del（SAS）無法由使用者模式合成；此情況改為開啟 Task Manager（等價的替代行為）。
        /// </summary>
        private void OnKeyClick(KeyModel? key)
        {
            if (key == null) return;

            // Mode 鍵行為
            if (key.VkCode == MODE_KEY_CODE)
            {
                // 若 Fn 開啟，Mode 僅切換本地預覽（不改系統 IME）
                if (_virtualFnToggle)
                {
                    _localPreviewToggle = !_localPreviewToggle;
                    _temporaryEnglishMode = false;
                    _temporaryEnglishFirstUpperSent = false;
                    _temporaryShiftInjected = false;
                    _virtualShiftToggle = false;
                    UpdateDisplay();
                    return;
                }

                // 注入 Shift down/up 切換系統 IME（模擬使用者按下 Shift）
                _suppressInjectedShiftCount = 2;
                keybd_event(0x10, 0, 0, 0);
                Thread.Sleep(5);
                keybd_event(0x10, 0, 2, 0);

                // 立即更新本地視覺預覽並在短時間內忽略系統 IME 回報
                _isZhuyinMode = !_isZhuyinMode;
                _ignoreImeSyncUntil = DateTime.UtcNow.AddMilliseconds(300);

                // 清除臨時與虛擬狀態
                _temporaryEnglishMode = false;
                _temporaryEnglishFirstUpperSent = false;
                _temporaryShiftInjected = false;
                _virtualShiftToggle = false;

                UpdateDisplay();
                return;
            }

            // Fn/Ctrl/Alt/Win 虛擬切換
            if (key.VkCode == FN_KEY_CODE) { _virtualFnToggle = !_virtualFnToggle; UpdateDisplay(); return; }
            if (key.VkCode == 0x11) { _virtualCtrlToggle = !_virtualCtrlToggle; UpdateDisplay(); return; }
            if (key.VkCode == 0x12)
            {
                // Toggle virtual Alt. When turning on, inject Alt down and keep it until toggled off.
                _virtualAltToggle = !_virtualAltToggle;
                if (_virtualAltToggle)
                {
                    keybd_event(0x12, 0, 0, 0); // Alt down
                }
                else
                {
                    keybd_event(0x12, 0, 2, 0); // Alt up
                }
                UpdateDisplay();
                return;
            }
            if (key.VkCode == 0x5B) { _virtualWinToggle = !_virtualWinToggle; UpdateDisplay(); return; }

            // Shift 行為：在注音模式下進入臨時英文模式（優先），否則切換虛擬 Shift
            if (key.VkCode == 0x10 && _isZhuyinMode)
            {
                _temporaryEnglishMode = true;
                _temporaryEnglishFirstUpperSent = false;
                UpdateDisplay();
                return;
            }
            if (key.VkCode == 0x10) { _virtualShiftToggle = !_virtualShiftToggle; UpdateDisplay(); return; }

            // Fn 模式下的 Tab (在 Fn 位置顯示為 Ctrl+Alt+Del)：直接開啟 Task Manager 作為 SAS 的替代
            if (_virtualFnToggle && key.VkCode == 0x09)
            {
                // 顯示自訂的安全選單（鎖定 / 切換使用者 / 登出 / 變更密碼 / 工作管理員 / 取消）
                ShowSecurityMenu();

                // 清除虛擬 modifiers（臨時英文模式除外）
                // 注意：Alt 要作為鎖定鍵，不在此自動清除，必須再按 Alt 才會釋放
                if (!_temporaryEnglishMode)
                {
                    _virtualShiftToggle = false;
                    _virtualCtrlToggle = false;
                    _virtualWinToggle = false;
                    // _virtualAltToggle 保持鎖定狀態
                }
                UpdateDisplay();
                return;
            }

            // 處理一般按鍵的發送，考慮 Fn 映射與目前的 modifiers
            byte sendVk = key.VkCode;
            if (_virtualFnToggle && FnSendMap.TryGetValue(key.VkCode, out byte targetVk)) sendVk = targetVk;

            // 若 virtual Alt 為鎖定（sticky），且使用者按的是上排數字，轉成 NumPad 對應鍵以支援 Alt+Numpad 輸入序列
            if (_virtualAltToggle && key.VkCode >= 0x30 && key.VkCode <= 0x39)
            {
                sendVk = (byte)(0x60 + (key.VkCode - 0x30)); // VK_NUMPAD0..VK_NUMPAD9 = 0x60..0x69
            }

            // 特殊處理：若為 Fn 下的 Delete（實際 sendVk == 0x2E）且同時有 Ctrl 與 Alt（SAS-like）
            // Windows 的安全注意鍵 Ctrl+Alt+Del (SAS) 無法被使用者模式合成，故改為開啟 Task Manager 作為替代行為。
            if (sendVk == 0x2E && _isCtrlActive && _isAltActive)
            {
                try
                {
                    // 使用 taskmgr 作為替代，等同於使用者開啟 Task Manager（也可使用 Ctrl+Shift+Esc）
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("taskmgr") { UseShellExecute = true });
                }
                catch
                {
                    // 若失敗則 fallback 為直接注入 Delete（但可能無法達成 SAS）
                    keybd_event(sendVk, 0, 0, 0);
                    keybd_event(sendVk, 0, 2, 0);
                }

                // 清除虛擬 modifiers（除非在臨時英文模式）
                // 不自動清除 Alt，Alt 為鎖定鍵，需再次按 Alt 才會釋放
                if (!_temporaryEnglishMode)
                {
                    _virtualShiftToggle = false;
                    _virtualCtrlToggle = false;
                    _virtualWinToggle = false;
                    // _virtualAltToggle 保持
                }
                UpdateDisplay();
                return;
            }

            var mods = new List<byte>();
            if (_isCtrlActive) { keybd_event(0x11, 0, 0, 0); mods.Add(0x11); }
            // If virtual Alt is toggled-on, Alt is already injected and should not be injected again for this key
            if (_isAltActive && !_virtualAltToggle) { keybd_event(0x12, 0, 0, 0); mods.Add(0x12); }
            if (_isWinActive) { keybd_event(0x5B, 0, 0, 0); mods.Add(0x5B); }

            bool isAlpha = sendVk >= 0x41 && sendVk <= 0x5A;

            // 臨時英文模式下，第一個字母以 Shift+letter 注入（若尚未送出），其後字母以單字母送出
            if (_temporaryEnglishMode && isAlpha)
            {
                if (!_temporaryEnglishFirstUpperSent)
                {
                    keybd_event(0x10, 0, 0, 0); // shift down
                    keybd_event(sendVk, 0, 0, 0);
                    keybd_event(sendVk, 0, 2, 0);
                    keybd_event(0x10, 0, 2, 0); // shift up
                    _temporaryEnglishFirstUpperSent = true;
                }
                else
                {
                    keybd_event(sendVk, 0, 0, 0);
                    keybd_event(sendVk, 0, 2, 0);
                }
            }
            else
            {
                keybd_event(sendVk, 0, 0, 0);
                keybd_event(sendVk, 0, 2, 0);
            }

            // 釋放先前按下的虛擬 modifiers（以 LIFO 釋放順序）
            for (int i = mods.Count - 1; i >= 0; i--) keybd_event(mods[i], 0, 2, 0);

            // 若送出的不是 modifier 則預設清除虛擬 modifiers（臨時英文模式除外）
            // Alt 為鎖定鍵，不會在此自動清除
            if (sendVk != 0x11 && sendVk != 0x12 && sendVk != 0x10 && sendVk != 0x5B)
            {
                if (!_temporaryEnglishMode)
                {
                    _virtualShiftToggle = false;
                    _virtualCtrlToggle = false;
                    _virtualWinToggle = false;
                    // _virtualAltToggle 保持
                }
            }

            // 按下 Enter 且在臨時英文模式下結束該模式（不改變系統 IME）
            if (_temporaryEnglishMode && key.VkCode == 0x0D)
            {
                _temporaryEnglishMode = false;
                _temporaryEnglishFirstUpperSent = false;
            }

            UpdateDisplay();
        }

        /// <summary>
        /// 根據目前系統與虛擬狀態更新所有按鍵的顯示名稱與顏色：
        /// - 判斷大寫/小寫（CapsLock XOR Shift）。
        /// - 支援符號顯示（Shift）與本地預覽切換。
        /// - 臨時英文模式會以特別色彩指示第一個字母為大寫預覽。
        /// - Fn 模式會改變某些鍵的顯示文字與顏色。
        /// </summary>
        private void UpdateDisplay()
        {
            bool upper = _isCapsLockActive ^ _isShiftActive;
            bool symbols = _isShiftActive;

            // displayZhuyin 決定底層按鍵是否顯示注音（視 local preview 與實際 IME 狀態）
            bool displayZhuyin = _localPreviewToggle ? !_isZhuyinMode : _isZhuyinMode;
            ModeIndicator = displayZhuyin ? "En" : "ㄅ";
            IndicatorColor = displayZhuyin ? "Orange" : "White";

            if (_temporaryEnglishMode)
            {
                ModeIndicator = "En";
                IndicatorColor = "Cyan";
            }

            foreach (var row in KeyRows)
            {
                foreach (var k in row)
                {
                    // Mode 鍵顯示目前的 ModeIndicator
                    if (k.VkCode == MODE_KEY_CODE)
                    {
                        k.DisplayName = ModeIndicator;
                        k.TextColor = IndicatorColor;
                        continue;
                    }

                    // Fn 模式改變顯示 (eg. 1 -> F1)
                    if (_virtualFnToggle && FnDisplayMap.TryGetValue(k.VkCode, out string? fnLabel))
                    {
                        if (!string.IsNullOrEmpty(fnLabel)) { k.DisplayName = fnLabel!; k.TextColor = "LightBlue"; continue; }
                    }

                    // Fn 按鍵本身的顯示
                    if (k.VkCode == FN_KEY_CODE)
                    {
                        k.DisplayName = "⌨";
                        k.TextColor = _virtualFnToggle ? "Cyan" : "White";
                        continue;
                    }

                    bool isAlpha = k.VkCode >= 0x41 && k.VkCode <= 0x5A;

                    // 若系統為中文（注音）且不在 symbols (Shift) 且該鍵有注音字串，顯示注音
                    if (_isZhuyinMode && !symbols && !string.IsNullOrEmpty(k.Zhuyin) && !_temporaryEnglishMode)
                    {
                        k.DisplayName = k.Zhuyin;
                        k.TextColor = "Orange";
                    }
                    else
                    {
                        // 臨時英文模式對字母的特殊顯示：第一個字母大寫預覽
                        if (_temporaryEnglishMode && isAlpha)
                        {
                            if (!_temporaryEnglishFirstUpperSent)
                            {
                                k.DisplayName = k.EnglishUpper;
                                k.TextColor = "LightBlue";
                            }
                            else
                            {
                                k.DisplayName = k.English;
                                k.TextColor = "White";
                            }
                        }
                        else
                        {
                            k.DisplayName = (isAlpha ? upper : symbols) ? k.EnglishUpper : k.English;
                            k.TextColor = (isAlpha ? upper : symbols) ? "LightBlue" : "White";
                        }
                    }

                    // 個別 modifier 鍵的顏色高亮顯示
                    if (k.VkCode == 0x14) k.TextColor = _isCapsLockActive ? "Cyan" : "White";
                    if (k.VkCode == 0x10) k.TextColor = _isShiftActive ? "Cyan" : "White";
                    if (k.VkCode == 0x11) k.TextColor = _isCtrlActive ? "Cyan" : "White";
                    if (k.VkCode == 0x12) k.TextColor = _isAltActive ? "Cyan" : "White";
                    if (k.VkCode == 0x5B) k.TextColor = _isWinActive ? "Cyan" : "White";
                }
            }
        }

        /// <summary>
        /// 建立鍵盤的資料模型 (KeyRows) 與每鍵屬性（英文、英文字母大寫、注音、VkCode、寬度）。
        /// 此方法只建立視覺與發送對應的按鍵資訊，不會改變系統狀態。
        /// </summary>
        private void SetupKeyboard()
        {
            var r1 = new ObservableCollection<KeyModel>();
            r1.Add(new KeyModel { English = "`", EnglishUpper = "~", Zhuyin = "", VkCode = 0xC0, Width = 65 });
            r1.Add(new KeyModel { English = "1", EnglishUpper = "!", Zhuyin = "ㄅ", VkCode = 0x31 });
            r1.Add(new KeyModel { English = "2", EnglishUpper = "@", Zhuyin = "ㄉ", VkCode = 0x32 });
            r1.Add(new KeyModel { English = "3", EnglishUpper = "#", Zhuyin = "ˇ", VkCode = 0x33 });
            r1.Add(new KeyModel { English = "4", EnglishUpper = "$", Zhuyin = "ˋ", VkCode = 0x34 });
            r1.Add(new KeyModel { English = "5", EnglishUpper = "%", Zhuyin = "ㄓ", VkCode = 0x35 });
            r1.Add(new KeyModel { English = "6", EnglishUpper = "^", Zhuyin = "ˊ", VkCode = 0x36 });
            r1.Add(new KeyModel { English = "7", EnglishUpper = "&", Zhuyin = "˙", VkCode = 0x37 });
            r1.Add(new KeyModel { English = "8", EnglishUpper = "*", Zhuyin = "ㄚ", VkCode = 0x38 });
            r1.Add(new KeyModel { English = "9", EnglishUpper = "(", Zhuyin = "ㄞ", VkCode = 0x39 });
            r1.Add(new KeyModel { English = "0", EnglishUpper = ")", Zhuyin = "ㄢ", VkCode = 0x30 });
            r1.Add(new KeyModel { English = "-", EnglishUpper = "_", Zhuyin = "ㄦ", VkCode = 0xBD });
            r1.Add(new KeyModel { English = "=", EnglishUpper = "+", Zhuyin = "", VkCode = 0xBB });
            r1.Add(new KeyModel { English = "⌫", EnglishUpper = "⌫", Zhuyin = "", VkCode = 0x08, Width = 95 });
            KeyRows.Add(r1);

            var r2 = new ObservableCollection<KeyModel>();
            r2.Add(new KeyModel { English = "⇥", EnglishUpper = "Tab", Zhuyin = "", VkCode = 0x09, Width = 95 });
            r2.Add(new KeyModel { English = "q", EnglishUpper = "Q", Zhuyin = "ㄆ", VkCode = 0x51 });
            r2.Add(new KeyModel { English = "w", EnglishUpper = "W", Zhuyin = "ㄊ", VkCode = 0x57 });
            r2.Add(new KeyModel { English = "e", EnglishUpper = "E", Zhuyin = "ㄍ", VkCode = 0x45 });
            r2.Add(new KeyModel { English = "r", EnglishUpper = "R", Zhuyin = "ㄐ", VkCode = 0x52 });
            r2.Add(new KeyModel { English = "t", EnglishUpper = "T", Zhuyin = "ㄔ", VkCode = 0x54 });
            r2.Add(new KeyModel { English = "y", EnglishUpper = "Y", Zhuyin = "ㄗ", VkCode = 0x59 });
            r2.Add(new KeyModel { English = "u", EnglishUpper = "U", Zhuyin = "ㄧ", VkCode = 0x55 });
            r2.Add(new KeyModel { English = "i", EnglishUpper = "I", Zhuyin = "ㄛ", VkCode = 0x49 });
            r2.Add(new KeyModel { English = "o", EnglishUpper = "O", Zhuyin = "ㄟ", VkCode = 0x4F });
            r2.Add(new KeyModel { English = "p", EnglishUpper = "P", Zhuyin = "ㄣ", VkCode = 0x50 });
            r2.Add(new KeyModel { English = "[", EnglishUpper = "{", Zhuyin = "", VkCode = 0xDB });
            r2.Add(new KeyModel { English = "]", EnglishUpper = "}", Zhuyin = "", VkCode = 0xDD });
            r2.Add(new KeyModel { English = "\\", EnglishUpper = "|", Zhuyin = "", VkCode = 0xDC, Width = 65 });
            KeyRows.Add(r2);

            var r3 = new ObservableCollection<KeyModel>();
            r3.Add(new KeyModel { English = "⇪", EnglishUpper = "Caps", Zhuyin = "", VkCode = 0x14, Width = 133 });
            r3.Add(new KeyModel { English = "a", EnglishUpper = "A", Zhuyin = "ㄇ", VkCode = 0x41 });
            r3.Add(new KeyModel { English = "s", EnglishUpper = "S", Zhuyin = "ㄋ", VkCode = 0x53 });
            r3.Add(new KeyModel { English = "d", EnglishUpper = "D", Zhuyin = "ㄎ", VkCode = 0x44 });
            r3.Add(new KeyModel { English = "f", EnglishUpper = "F", Zhuyin = "ㄙ", VkCode = 0x46 });
            r3.Add(new KeyModel { English = "g", EnglishUpper = "G", Zhuyin = "ㄕ", VkCode = 0x47 });
            r3.Add(new KeyModel { English = "h", EnglishUpper = "H", Zhuyin = "ㄘ", VkCode = 0x48 });
            r3.Add(new KeyModel { English = "j", EnglishUpper = "J", Zhuyin = "ㄨ", VkCode = 0x4A });
            r3.Add(new KeyModel { English = "k", EnglishUpper = "K", Zhuyin = "ㄜ", VkCode = 0x4B });
            r3.Add(new KeyModel { English = "l", EnglishUpper = "L", Zhuyin = "ㄠ", VkCode = 0x4C });
            r3.Add(new KeyModel { English = ";", EnglishUpper = ":", Zhuyin = "ㄤ", VkCode = 0xBA });
            r3.Add(new KeyModel { English = "'", EnglishUpper = "\"", Zhuyin = "ㄦ", VkCode = 0xDE });
            r3.Add(new KeyModel { English = "⏎", EnglishUpper = "Enter", Zhuyin = "送出", VkCode = 0x0D, Width = 98 });
            KeyRows.Add(r3);

            var r4 = new ObservableCollection<KeyModel>();
            r4.Add(new KeyModel { English = "⇧", EnglishUpper = "Shift", Zhuyin = "", VkCode = 0x10, Width = 166 });
            r4.Add(new KeyModel { English = "z", EnglishUpper = "Z", Zhuyin = "ㄈ", VkCode = 0x5A });
            r4.Add(new KeyModel { English = "x", EnglishUpper = "X", Zhuyin = "ㄌ", VkCode = 0x58 });
            r4.Add(new KeyModel { English = "c", EnglishUpper = "C", Zhuyin = "ㄏ", VkCode = 0x43 });
            r4.Add(new KeyModel { English = "v", EnglishUpper = "V", Zhuyin = "ㄒ", VkCode = 0x56 });
            r4.Add(new KeyModel { English = "b", EnglishUpper = "B", Zhuyin = "ㄖ", VkCode = 0x42 });
            r4.Add(new KeyModel { English = "n", EnglishUpper = "N", Zhuyin = "ㄙ", VkCode = 0x4E });
            r4.Add(new KeyModel { English = "m", EnglishUpper = "M", Zhuyin = "ㄩ", VkCode = 0x4D });
            r4.Add(new KeyModel { English = ",", EnglishUpper = "<", Zhuyin = "ㄝ", VkCode = 0xBC });
            r4.Add(new KeyModel { English = ".", EnglishUpper = ">", Zhuyin = "ㄡ", VkCode = 0xBE });
            r4.Add(new KeyModel { English = "/", EnglishUpper = "?", Zhuyin = "ㄥ", VkCode = 0xBF });
            r4.Add(new KeyModel { English = "↑", EnglishUpper = "↑", Zhuyin = "", VkCode = 0x26 });
            r4.Add(new KeyModel { English = "⌨", EnglishUpper = "Fn", Zhuyin = "", VkCode = FN_KEY_CODE });
            KeyRows.Add(r4);

            var r5 = new ObservableCollection<KeyModel>();
            r5.Add(new KeyModel { English = "⌃", EnglishUpper = "Ctrl", Zhuyin = "", VkCode = 0x11 });
            r5.Add(new KeyModel { English = "⊞", EnglishUpper = "Win", Zhuyin = "", VkCode = 0x5B });
            r5.Add(new KeyModel { English = "⌥", EnglishUpper = "Alt", Zhuyin = "", VkCode = 0x12 });
            r5.Add(new KeyModel { English = "⎵", EnglishUpper = "Space", Zhuyin = "空白鍵", VkCode = 0x20, Width = 512 });
            r5.Add(new KeyModel { English = "Mode", EnglishUpper = "Mode", Zhuyin = "", VkCode = MODE_KEY_CODE });
            r5.Add(new KeyModel { English = "←", EnglishUpper = "←", Zhuyin = "", VkCode = 0x25 });
            r5.Add(new KeyModel { English = "↓", EnglishUpper = "↓", Zhuyin = "", VkCode = 0x28 });
            r5.Add(new KeyModel { English = "→", EnglishUpper = "→", Zhuyin = "", VkCode = 0x27 });
            KeyRows.Add(r5);

            UpdateDisplay();
        }

        // UI 事件：縮小、退出、拖曳與調整尺寸
        private void Minimize_Click(object sender, RoutedEventArgs e) => this.Hide();
        private void Exit_Click(object sender, RoutedEventArgs e) => System.Windows.Application.Current.Shutdown();
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) { if (!_isResizing) this.DragMove(); }
        private void Resize_Init(object sender, MouseButtonEventArgs e) { _isResizing = true; Mouse.Capture((UIElement)sender); }
        private void Resize_End(object sender, MouseButtonEventArgs e) { _isResizing = false; Mouse.Capture(null); }
        private void Resize_Move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isResizing)
            {
                System.Windows.Point p = e.GetPosition(this);
                if (p.X > 200) this.Width = p.X;
                if (p.Y > 150) this.Height = p.Y;
            }
        }

        /// <summary>
        /// 視窗關閉時解除掛鉤並釋放系統托盤圖示資源。
        /// </summary>
        protected override void OnClosed(EventArgs e) { UnhookWindowsHookEx(_hookID); _notifyIcon.Dispose(); base.OnClosed(e); }

        // Ensure virtual Alt is released if still toggled when window closes
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_virtualAltToggle)
            {
                keybd_event(0x12, 0, 2, 0); // Alt up
                _virtualAltToggle = false;
            }
            base.OnClosing(e);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        /// <summary>
        /// Slider 點擊軌道時直接將滑桿設到點擊位置（回退處理，當 IsMoveToPointEnabled 不可用時）
        /// </summary>
        private void OpacitySlider_PreviewMouseLeftButtonDown(object? sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is not System.Windows.Controls.Slider slider) return;
            var pt = e.GetPosition(slider);
            double ratio = pt.X / slider.ActualWidth;
            if (ratio < 0) ratio = 0; if (ratio > 1) ratio = 1;
            double newValue = slider.Minimum + (slider.Maximum - slider.Minimum) * ratio;
            slider.Value = newValue;
            e.Handled = true;
        }

        // 顯示自訂的安全選單（模擬 Ctrl+Alt+Del 的選項）
        private void ShowSecurityMenu()
        {
            // 在 UI 執行緒顯示 WPF ContextMenu
            this.Dispatcher.Invoke(() =>
            {
                var menu = new System.Windows.Controls.ContextMenu();

                var miLock = new System.Windows.Controls.MenuItem { Header = "鎖定" };
                miLock.Click += (s, e) => { LockWorkStation(); };

                var miSwitch = new System.Windows.Controls.MenuItem { Header = "切換使用者" };
                miSwitch.Click += (s, e) => { StartSwitchUser(); };

                var miSignOut = new System.Windows.Controls.MenuItem { Header = "登出" };
                miSignOut.Click += (s, e) => { ExitWindowsEx(0x00000000u /* EWX_LOGOFF */, 0); };

                var miChangePwd = new System.Windows.Controls.MenuItem { Header = "變更密碼" };
                miChangePwd.Click += (s, e) => { ShowChangePassword(); };

                var miTaskMgr = new System.Windows.Controls.MenuItem { Header = "工作管理員" };
                miTaskMgr.Click += (s, e) => { TryStartTaskManager(); };

                var miCancel = new System.Windows.Controls.MenuItem { Header = "取消" };
                miCancel.Click += (s, e) => { /* no-op */ };

                menu.Items.Add(miLock);
                menu.Items.Add(miSwitch);
                menu.Items.Add(miSignOut);
                menu.Items.Add(miChangePwd);
                menu.Items.Add(new System.Windows.Controls.Separator());
                menu.Items.Add(miTaskMgr);
                menu.Items.Add(miCancel);

                // 在視窗右上角或按鍵附近顯示：這裡顯示在視窗中心上方
                menu.PlacementTarget = this;
                menu.Placement = System.Windows.Controls.Primitives.PlacementMode.Center;
                menu.IsOpen = true;
            });
        }

        private void TryStartTaskManager()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("taskmgr") { UseShellExecute = true });
            }
            catch
            {
                // fallback: inject Ctrl+Shift+Esc
                keybd_event(0x11, 0, 0, 0); // Ctrl down
                keybd_event(0x10, 0, 0, 0); // Shift down
                keybd_event(0x1B, 0, 0, 0); // Esc down
                keybd_event(0x1B, 0, 2, 0); // Esc up
                keybd_event(0x10, 0, 2, 0); // Shift up
                keybd_event(0x11, 0, 2, 0); // Ctrl up
            }
        }

        private void StartSwitchUser()
        {
            try
            {
                // tsdiscon 會在遠端桌面情況下可用；作為一般使用者嘗試呼叫鎖定畫面，使用tsdiscon或tscon可能需要權限。
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("tsdiscon") { UseShellExecute = true });
            }
            catch
            {
                // 無法使用 tsdiscon，嘗試顯示切換使用者對話或 fallback 為 LockWorkStation
                LockWorkStation();
            }
        }

        private void ShowChangePassword()
        {
            try
            {
                // 直接呼叫變更密碼介面並非簡單；使用控制台帳戶頁面作為 fallback
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("control", "/name Microsoft.UserAccounts") { UseShellExecute = true });
            }
            catch
            {
                // fallback: no-op
            }
        }
    }

    /// <summary>
    /// KeyModel：代表單一按鍵的顯示與邏輯屬性，可透過 INotifyPropertyChanged 更新 UI。
    /// 屬性：
    /// - English / EnglishUpper: 英文字母小寫/大寫或符號。
    /// - Zhuyin: 注音顯示字串（若該按鍵在注音模式下顯示）。
    /// - VkCode: 與該鍵對應的虛擬鍵碼（若為自定義 Mode/Fn，會使用非標準 VK 值）。
    /// - Width: 在 UI 中的寬度 (像素)。
    /// - DisplayName / TextColor / Background / IsPressed: 由 MainWindow.UpdateDisplay 與 HookCallback 更新以反映視覺狀態。
    /// </summary>
    public class KeyModel : INotifyPropertyChanged
    {
        public string English { get; set; } = "";
        public string EnglishUpper { get; set; } = "";
        public string Zhuyin { get; set; } = "";
        public byte VkCode { get; set; }
        public double Width { get; set; } = 65;

        private string _displayName = "";
        public string DisplayName { get { return _displayName; } set { _displayName = value; OnPropertyChanged("DisplayName"); } }

        private string _textColor = "White";
        public string TextColor { get { return _textColor; } set { _textColor = value; OnPropertyChanged("TextColor"); } }

        private bool _isPressed = false;
        public bool IsPressed { get { return _isPressed; } set { _isPressed = value; OnPropertyChanged("IsPressed"); OnPropertyChanged("Background"); } }

        public string Background => _isPressed ? "#666666" : "#333333";

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 簡易的 RelayCommand 實作，用於將按鍵點擊綁定到 ViewModel (MainWindow) 中的處理器。
    /// </summary>
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T?> _execute;
        public RelayCommand(Action<T?> execute) => _execute = execute;
        public bool CanExecute(object? parameter) => true;
        public void Execute(object? parameter) => _execute((T?)parameter);
        public event EventHandler? CanExecuteChanged;
        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}