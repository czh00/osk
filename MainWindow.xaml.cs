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
using System.Windows.Media.Animation;
using System.Reflection;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;

namespace OSK
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        #region Win32 / IMM API 宣告
        private InputDetector? _inputDetector;
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        // --- SendInput 相關結構與 API ---
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion u;
            public static int Size => Marshal.SizeOf(typeof(INPUT));
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const int INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_SCANCODE = 0x0008;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        // ---------------------------------

        [DllImport("user32.dll")] private static extern uint RegisterWindowMessage(string lpString);
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        // 系統操作 API
        [DllImport("user32.dll", SetLastError = true)] private static extern bool LockWorkStation();
        [DllImport("user32.dll", SetLastError = true)] private static extern bool ExitWindowsEx(uint uFlags, uint dwReason);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        private const int HWND_BROADCAST = 0xFFFF;

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

        // IMM32 API
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetContext(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmGetConversionStatus(IntPtr hIMC, out uint lpdwConversion, out uint lpdwSentence);
        [DllImport("imm32.dll")] private static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);
        [DllImport("imm32.dll")] private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);
        [DllImport("imm32.dll")] private static extern bool ImmSetConversionStatus(IntPtr hIMC, uint dwConversion, uint dwSentence);
        [DllImport("imm32.dll")] private static extern bool ImmGetOpenStatus(IntPtr hIMC);
        [DllImport("imm32.dll")] private static extern bool ImmSetOpenStatus(IntPtr hIMC, bool fOpen);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint WM_IME_CONTROL = 0x0283;
        private const int IMC_GETCONVERSIONMODE = 0x0001;
        private const int IMC_GETOPENSTATUS = 0x0005;

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPOS
        {
            public IntPtr hwnd;
            public IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public uint flags;
        }

        [DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, int minimumWorkingSetSize, int maximumWorkingSetSize);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        
        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(System.Drawing.Point pt, uint dwFlags);

        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        private const int MONITOR_DEFAULTTONEAREST = 2;
        private const int MDT_EFFECTIVE_DPI = 0;

        private const int DWMWA_TRANSITIONS_FORCEDISABLED = 3;
        private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;
        private const int DWMWCP_ROUND = 2;

        #endregion

        #region 常數與委派
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WM_WINDOWPOSCHANGING = 0x0046;
        private const int WM_NCHITTEST = 0x0084;
        private const int HTCLIENT = 1;

        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint IME_CMODE_NATIVE = 0x0001;
        private const byte MODE_KEY_CODE = 0xFF;
        private const byte FN_KEY_CODE = 0xFE;
        private const byte CAD_KEY_CODE = 0xFD;
        #endregion

        #region 欄位
        private uint _msgShowOsk;
        private System.Windows.Forms.NotifyIcon? _notifyIcon;
        private static Mutex? _mutex;
        private static readonly Random _rnd = new();

        public ObservableCollection<ObservableCollection<KeyModel>> KeyRows { get; set; } = new();
        public ICommand? KeyCommand { get; set; }
        public ICommand? ToggleThemeCommand { get; set; }
        public ICommand? TogglePinCommand { get; set; }
        public ICommand? ToggleDragFxCommand { get; set; }
        public ICommand? ToggleLayoutCommand { get; set; }
        public ICommand? ToggleFullLayoutCommand { get; set; }
        public ICommand? ToggleEditModeCommand { get; set; }

        private bool _isPinned = true;
        public string PinIcon { get { return _isPinned ? "📍" : "📌"; } }

        private bool _isDragFxEnabled = true;
        public string DragFxIcon { get { return _isDragFxEnabled ? "❄️" : "🌊"; } }

        private bool _isEditMode = false;
        public bool IsEditMode { get { return _isEditMode; } set { _isEditMode = value; OnPropertyChanged("IsEditMode"); OnPropertyChanged("SettingsBtnColor"); } }
        public string SettingsBtnColor { get { return _isEditMode ? _themeActiveColor : UiTextColor; } }

        // Drag and Drop state
        private bool _isManualDragging = false;
        private System.Windows.Point _dragStartPoint;
        private KeyModel? _draggedKey;
        private System.Windows.Controls.Border? _draggedBorder;
        private System.Windows.Shapes.Rectangle? _dragGhost;
        private System.Windows.Point _ghostOffset;

        private bool _isDynamicLayout = true;
        public bool IsDynamicLayout { get { return _isDynamicLayout; } set { _isDynamicLayout = value; OnPropertyChanged("IsDynamicLayout"); } }

        private string _layoutIcon = "🗔";
        public string LayoutIcon { get { return _layoutIcon; } set { _layoutIcon = value; OnPropertyChanged("LayoutIcon"); } }

        private bool _isFullLayout = false;
        public bool IsFullLayout { get { return _isFullLayout; } set { _isFullLayout = value; OnPropertyChanged("IsFullLayout"); OnPropertyChanged("KeyboardAlignment"); } }

        public string KeyboardAlignment => "Left";

        private string _fullLayoutIcon = "⌨"; // 預設為精簡，圖示應顯示「切換至全鍵盤」
        public string FullLayoutIcon { get { return _fullLayoutIcon; } set { _fullLayoutIcon = value; OnPropertyChanged("FullLayoutIcon"); } }

        private bool _isZhuyinMode = false;
        private bool _isShiftActive = false;
        private bool _isCapsLockActive = false;
        private bool _isCtrlActive = false;
        private bool _isAltActive = false;
        private bool _isWinActive = false;

        private bool _lastPhysicalShiftDown = false;
        private DateTime _physShiftDownTime = DateTime.MinValue;
        private bool _isNumLockActive = false;
        private bool _isTouchDraggingWindow = false;
        private System.Windows.Point _touchDragStartPoint;

        private string _currentAppInfo = "OSK (讀取中...)";
        public string CurrentAppInfo { get { return _currentAppInfo; } set { _currentAppInfo = value; OnPropertyChanged("CurrentAppInfo"); } }
        private string _lastLoggedAppInfo = "";
        private uint _lastPid = 0;
        private string? _lastProcessName = null;
        private int _clipboardCheckCounter = 0;
        private string _lastClipboardText = "";

        // 本地虛擬 modifier
        private bool _virtualShiftToggle = false;
        private byte _activeShiftVk = 0x10;
        private bool _virtualCtrlToggle = false;
        private byte _activeCtrlVk = 0x11;
        private bool _virtualAltToggle = false;
        private byte _activeAltVk = 0x12;
        private bool _virtualWinToggle = false;
        private byte _activeWinVk = 0x5B;
        private bool _virtualFnToggle = false;

        private bool _localPreviewToggle = false;

        // 臨時英文模式
        private bool _temporaryEnglishMode = false;
        private DateTime _ignoreImeSyncUntil = DateTime.MinValue;

        private string _modeIndicator = "En";
        public string ModeIndicator { get { return _modeIndicator; } set { _modeIndicator = value; OnPropertyChanged("ModeIndicator"); } }

        private string _indicatorColor = "White";
        public string IndicatorColor { get { return _indicatorColor; } set { _indicatorColor = value; OnPropertyChanged("IndicatorColor"); } }

        // --- 主題相關屬性 ---
        private bool _isDarkMode = true;

        // 主題圖示 (🌙/☀)
        private string _themeIcon = "🌙";
        public string ThemeIcon { get { return _themeIcon; } set { _themeIcon = value; OnPropertyChanged("ThemeIcon"); } }

        // 視窗背景
        private string _windowBackground = "#1E1E1E";
        public string WindowBackground { get { return _windowBackground; } set { _windowBackground = value; OnPropertyChanged("WindowBackground"); } }

        // UI 介面文字/圖示顏色 (用於控制列按鈕、標籤)
        private string _uiTextColor = "White";
        public string UiTextColor { get { return _uiTextColor; } set { _uiTextColor = value; OnPropertyChanged("UiTextColor"); } }

        // 用於切換動態版面文字顏色
        private string _themeTextColor = "White";
        private string _themeActiveColor = "Cyan";
        private string _themeSubColor = "#1E90FF";
        private string _themeNumColor = "LightSkyBlue";

        // 縮放手柄顏色 (Resize Grip)
        private string _resizeGripColor = "#888888";
        public string ResizeGripColor { get { return _resizeGripColor; } set { _resizeGripColor = value; OnPropertyChanged("ResizeGripColor"); } }

        // 視窗控制按鈕背景 (最小化按鈕)
        private string _controlBtnBackground = "#333333";
        public string ControlBtnBackground { get { return _controlBtnBackground; } set { _controlBtnBackground = value; OnPropertyChanged("ControlBtnBackground"); } }


        // Mode 鍵長按處理
        private DispatcherTimer _modeKeyTimer = new DispatcherTimer();
        private bool _modeKeyLongPressHandled = false;

        // 視覺同步 (使用 CompositionTarget.Rendering 取代 DispatcherTimer 以對齊 vsync)
        private DispatcherTimer _memoryTrimTimer = new DispatcherTimer();
        private DispatcherTimer _imeTimer = new DispatcherTimer();
        private IntPtr _hHook = IntPtr.Zero;

        // Fn 模式對照表
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
            [0x09] = "⌃⌥⌦",
            [0x08] = "⌦"
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
            [0x09] = 0x2E,
            [0x08] = 0x2E
        };

        private bool _isResizing = false;
        private bool _needsInitialCentering = false;

        private string _iniFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "osk.ini");

        // 視窗隨動追蹤
        private double _lastWindowLeft = double.NaN;
        private double _lastWindowTop = double.NaN;
        private DateTime _lastFrameTime = DateTime.Now;
        private bool _isCurrentlyDragging = false;
        private bool _isPhysicsActive = false; // 按鍵是否正在物理運動中 (包括拖曳後的歸位動畫)
        
        // [強化] 綜合所有繁忙狀態：物理動畫中、滑鼠拖動中、觸控拖動中、縮放中，以及系統 OSK 叫出中
        private bool _isSystemOskActive = false;
        public bool IsBusySuppressing => _isPhysicsActive || _isMouseDraggingWindow || _isTouchDraggingWindow || _isResizing || _isSystemOskActive;
        private bool _wasBusySuppressing = false; 
        public bool IsPhysicsActive => _isPhysicsActive;
        private int _lastDisplayStateHash = -1; 
        private double _accumulatedDeltaX = 0; // 位移累積器，捕捉極微小位移
        private double _accumulatedDeltaY = 0;

        // Virtual Screen Target Properties
        public double TargetLeft
        {
            get => SystemParameters.VirtualScreenLeft + Canvas.GetLeft(KeyboardContainer);
            set
            {
                Canvas.SetLeft(KeyboardContainer, value - SystemParameters.VirtualScreenLeft);
                // [優化] 拖曳中不進行磁碟寫入，移至 MouseUp/TouchUp 執行
            }
        }

        public double TargetTop
        {
            get => SystemParameters.VirtualScreenTop + Canvas.GetTop(KeyboardContainer);
            set
            {
                Canvas.SetTop(KeyboardContainer, value - SystemParameters.VirtualScreenTop);
                // [優化] 拖曳中不進行磁碟寫入，移至 MouseUp/TouchUp 執行
            }
        }

        public double TargetWidth
        {
            get => KeyboardContainer.Width;
            set => KeyboardContainer.Width = value;
        }

        public double TargetHeight
        {
            get => KeyboardContainer.Height;
            set => KeyboardContainer.Height = value;
        }

        #endregion

        // 將特定按鍵加入到 User32 SendInput 結構清單中 (高相容性版)
        private void AddKeyInput(List<INPUT> inputs, byte vk, bool keyUp)
        {
            ushort scanCode = (ushort)MapVirtualKey(vk, 0); // MAPVK_VK_TO_VSC = 0
            uint flags = (keyUp ? KEYEVENTF_KEYUP : 0);

            // 針對擴展鍵 (Extended Key) 加入旗標，確保 Windows 與遠端桌面 (Splashtop) 正確識別
            if (IsExtendedKey(vk)) flags |= KEYEVENTF_EXTENDEDKEY;

            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = vk,         // 同時保留虛擬鍵
                        wScan = scanCode, // 同時保留掃描碼 (提高 AHK 風格相容性)
                        dwFlags = flags,  // 預設不強制 KEYEVENTF_SCANCODE，讓系統自行處理
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
            inputs.Add(input);
        }

        private bool IsExtendedKey(byte vk)
        {
            // 常見擴展鍵：方向鍵、Home/End、PageUp/Down、Ins/Del、左右 Alt/Ctrl 等
            // 注意：NumLock (0x90) 通常不屬於擴展鍵，移除它以避免相容性問題
            return (vk >= 0x21 && vk <= 0x2E) || (vk >= 0x25 && vk <= 0x28) || 
                   (vk == 0xA1 || vk == 0xA3 || vk == 0xA5) || (vk == 0x5B || vk == 0x5C);
        }

        // 傳送單一按鍵的按下或放開事件
        private void SendSimulatedKey(byte vk, bool isKeyUp)
        {
            var list = new List<INPUT>();
            AddKeyInput(list, vk, isKeyUp);
            SendInput((uint)list.Count, list.ToArray(), Marshal.SizeOf(typeof(INPUT)));
        }

        // 清除所有虛擬修飾鍵的暫存狀態
        private void ResetVirtualModifiers()
        {
            if (_virtualShiftToggle) { SendSimulatedKey(_activeShiftVk, true); _virtualShiftToggle = false; }
            if (_virtualCtrlToggle) { SendSimulatedKey(_activeCtrlVk, true); _virtualCtrlToggle = false; }
            if (_virtualAltToggle) { SendSimulatedKey(_activeAltVk, true); _virtualAltToggle = false; }
            if (_virtualWinToggle) { SendSimulatedKey(_activeWinVk, true); _virtualWinToggle = false; }
        }

        public MainWindow()
        {
            _msgShowOsk = RegisterWindowMessage("WM_SHOW_OSK");

            bool createdNew;
            _mutex = new Mutex(true, "Global\\OSK_Unique_Mutex", out createdNew);

            if (!createdNew)
            {
                PostMessage((IntPtr)HWND_BROADCAST, _msgShowOsk, IntPtr.Zero, IntPtr.Zero);
                System.Windows.Application.Current.Shutdown();
                return;
            }

            InitializeComponent();
            this.DataContext = this;

            _modeKeyTimer.Interval = TimeSpan.FromSeconds(0.2);
            _modeKeyTimer.Tick += ModeKeyTimer_Tick;

            this.AddHandler(UIElement.PreviewMouseLeftButtonDownEvent, new MouseButtonEventHandler(OnGlobalPreviewMouseDown), true);
            this.AddHandler(UIElement.PreviewMouseLeftButtonUpEvent, new MouseButtonEventHandler(OnGlobalPreviewMouseUp), true);

            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Loaded += MainWindow_Loaded;

            var opacitySlider = this.FindName("OpacitySlider") as System.Windows.Controls.Slider;
            if (opacitySlider != null)
            {
                try { opacitySlider.IsMoveToPointEnabled = true; } catch { }
                opacitySlider.ValueChanged += (s, e) => { this.Opacity = e.NewValue; };
            }

            KeyCommand = new RelayCommand<KeyModel>(OnKeyClick);
            ToggleThemeCommand = new RelayCommand<object>(ToggleTheme);
            TogglePinCommand = new RelayCommand<object>(TogglePin);
            ToggleDragFxCommand = new RelayCommand<object>(ToggleDragFx);
            ToggleLayoutCommand = new RelayCommand<object>(ToggleLayout);
            ToggleFullLayoutCommand = new RelayCommand<object>(ToggleFullLayout);
            ToggleEditModeCommand = new RelayCommand<object>(ToggleEditMode);

            SetupKeyboard();

            // [修正] 確保在讀取 INI 前有預設尺寸，防止 NaN 產生
            TargetWidth = 1000;
            double ratio = IsFullLayout ? 0.23 : 0.31;
            TargetHeight = Math.Max(TargetWidth * ratio + 45, 120);

            // 如果沒有 INI 設定檔，首次啟動應預設置中並靠下
            if (!File.Exists(_iniFilePath))
            {
                _needsInitialCentering = true;
            }

            // 讀取 INI 設定 (包含位置與自訂按鍵排列)
            LoadSettings();

            KeyBoardItemsControl.ItemsSource = null;
            KeyBoardItemsControl.ItemsSource = KeyRows;

            System.Drawing.Icon? trayIcon = null;
            try
            {
                string exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName ?? "";
                if (!string.IsNullOrEmpty(exePath) && System.IO.File.Exists(exePath))
                {
                    trayIcon = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                }
            }
            catch { }

            if (trayIcon == null)
            {
                try
                {
                    var resourceUri = new Uri("pack://application:,,,/Assets/app.ico");
                    var streamInfo = System.Windows.Application.GetResourceStream(resourceUri);
                    if (streamInfo != null) trayIcon = new System.Drawing.Icon(streamInfo.Stream);
                }
                catch { }
            }
            if (trayIcon == null) trayIcon = SystemIcons.Application;

            _notifyIcon = new System.Windows.Forms.NotifyIcon { Icon = trayIcon, Visible = true, Text = "OSK" };
            _notifyIcon.MouseClick += (s, e) =>
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    ToggleVisibility();
                }
            };
            var menu = new System.Windows.Forms.ContextMenuStrip();
            menu.Items.Add("結束", null, (s, e) => System.Windows.Application.Current.Shutdown());
            _notifyIcon.ContextMenuStrip = menu;

            // 使用 CompositionTarget.Rendering 讓物理計算與 WPF vsync 對齊，消除頓挫感
            System.Windows.Media.CompositionTarget.Rendering += VisualSyncTimer_Tick;


            _memoryTrimTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
            _memoryTrimTimer.Tick += (s, e) => { if (!_isPhysicsActive) ReduceMemoryUsage(); };
            _memoryTrimTimer.Start();

            // IME 檢測移至獨立的慢速計時器，避免每幀都呼叫昂貴的 Win32 API
            _imeTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
            _imeTimer.Tick += (s, e) => { try { DetectImeStatus(); } catch { } };
            _imeTimer.Start();

            _inputDetector = new InputDetector(this);
            if (!_isPinned)
            {
                // [優化] 異步啟動偵測器，防止 UIAutomation 初始化阻塞主執行緒導致啟動緩慢
                Dispatcher.BeginInvoke(new Action(() => {
                    _inputDetector.Start();
                }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            }
        }

        #region 主題切換邏輯

        private void DetectSystemTheme()
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (key != null)
                    {
                        object? val = key.GetValue("AppsUseLightTheme");
                        if (val is int iVal)
                        {
                            bool isSystemLight = (iVal == 1);
                            ApplyTheme(!isSystemLight);
                            return;
                        }
                    }
                }
            }
            catch { }
            ApplyTheme(true);
        }

        private void ToggleTheme(object? parameter)
        {
            ApplyTheme(!_isDarkMode);
        }

        private void ToggleDragFx(object? parameter)
        {
            _isDragFxEnabled = !_isDragFxEnabled;
            OnPropertyChanged("DragFxIcon");
        }

        private void ApplyPinState(bool pinned)
        {
            _isPinned = pinned;
            OnPropertyChanged("PinIcon");
            this.Topmost = pinned;

            if (_inputDetector != null)
            {
                if (pinned)
                {
                    _inputDetector.Stop();
                    this.Show(); // 確保固定後可見
                }
                else
                {
                    _inputDetector.Start();
                }
            }
        }

        private void ApplyLayoutType(bool dynamic)
        {
            IsDynamicLayout = dynamic;
            LayoutIcon = dynamic ? "🗔" : "🗖";
            UpdateDisplay();
        }

        private void TogglePin(object? parameter)
        {
            ApplyPinState(!_isPinned);
        }

        private void ToggleLayout(object? parameter)
        {
            ApplyLayoutType(!IsDynamicLayout);
        }

        private void ToggleFullLayout(object? parameter)
        {
            IsFullLayout = !IsFullLayout;
            FullLayoutIcon = IsFullLayout ? "📱" : "⌨";

            KeyRows.Clear();
            if (IsFullLayout)
            {
                SetupFullKeyboard();
                double targetH = Math.Max(TargetWidth * 0.23 + 5, 80);
                TargetHeight = targetH;
            }
            else
            {
                SetupKeyboard();
                double targetH = Math.Max(TargetWidth * 0.31 + 5, 80);
                TargetHeight = targetH;
            }

            // 切換版面後嘗試套用自訂排列
            ApplyCustomLayout();

            UpdateDisplay();
        }

        private void ToggleEditMode(object? parameter)
        {
            IsEditMode = !IsEditMode;
            if (!IsEditMode)
            {
                SaveSettings(); // 離開編輯模式時儲存
            }
        }

        private void ApplyTheme(bool isDark)
        {
            _isDarkMode = isDark;

            if (_isDarkMode)
            {
                // 深色模式：最高不透明度為純黑
                ThemeIcon = "☀";
                WindowBackground = "#FF000000";

                // UI 介面顏色
                UiTextColor = "White";
                ResizeGripColor = "#AAAAAA";
                ControlBtnBackground = "#44333333";

                // 按鍵顏色
                _themeTextColor = "White";
                _themeActiveColor = "Cyan";
                _themeSubColor = "#1E90FF";
                _themeNumColor = "LightSkyBlue";
                IndicatorColor = "White";
            }
            else
            {
                // 淺色模式：最高不透明度為純白/淺灰
                ThemeIcon = "🌙";
                WindowBackground = "#FFF0F0F0";

                // UI 介面顏色
                UiTextColor = "#333333";
                ResizeGripColor = "#666666";
                ControlBtnBackground = "#44DDDDDD";

                // 按鍵顏色
                _themeTextColor = "#333333";
                _themeActiveColor = "#0078D7";
                _themeSubColor = "#0078D7";
                _themeNumColor = "#00A2E8";
                IndicatorColor = "#333333";
            }

            foreach (var row in KeyRows)
            {
                foreach (var k in row)
                {
                    if (_isDarkMode)
                        k.SetThemeColors("#22FFFFFF", "#44FFFFFF", "White", "#1E90FF", "#32CD32", "Orange", "LightSkyBlue");
                    else
                        k.SetThemeColors("#88FFFFFF", "#BBFFFFFF", "#333333", "#0078D7", "#2E8B57", "#D2691E", "#00A2E8");
                }
            }
            UpdateDisplay();
        }

        #endregion

        #region Mode 鍵長按處理邏輯

        private void OnGlobalPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (IsEditMode) return;
            if ((e.OriginalSource as FrameworkElement)?.DataContext is KeyModel key)
            {
                if (key.VkCode == MODE_KEY_CODE)
                {
                    _modeKeyLongPressHandled = false;
                    _modeKeyTimer.Start();
                }
            }
        }

        private void OnGlobalPreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            _modeKeyTimer.Stop();
        }

        private void ModeKeyTimer_Tick(object? sender, EventArgs e)
        {
            _modeKeyTimer.Stop();
            _modeKeyLongPressHandled = true;

            _isZhuyinMode = !_isZhuyinMode;

            _localPreviewToggle = false;
            _temporaryEnglishMode = false;
            ResetVirtualModifiers();

            UpdateDisplay();
        }

        #endregion

        private bool _isClamping = false;
        private void ClampToScreen()
        {
            if (_isClamping) return;
            _isClamping = true;
            try
            {
                var cursorPos = System.Windows.Forms.Cursor.Position;
                var screen = System.Windows.Forms.Screen.FromPoint(cursorPos);
                var wa = screen.WorkingArea;

                // [修正 V5] 徹底解決三螢幕 + 混合 DPI 下的回彈問題。
                // 邏輯單位在跨螢幕時不可靠，因此我們在「實體像素空間」進行邊界判定。
                
                // 1. 取得鍵盤相對於螢幕的「目前實體位置」
                // 我們直接從 KeyboardContainer 獲取其在全域螢幕上的實體矩形
                System.Windows.Point physStart = KeyboardContainer.PointToScreen(new System.Windows.Point(0, 0));
                System.Windows.Point physEnd = KeyboardContainer.PointToScreen(new System.Windows.Point(KeyboardContainer.ActualWidth, KeyboardContainer.ActualHeight));
                
                double physX = physStart.X;
                double physY = physStart.Y;
                double physW = physEnd.X - physStart.X;
                double physH = physEnd.Y - physStart.Y;

                double newPhysX = physX;
                double newPhysY = physY;

                // 2. 在物理空間進行 Clamp (對比 WorkingArea)
                if (newPhysX < wa.Left) newPhysX = wa.Left;
                if (newPhysX + physW > wa.Right) newPhysX = wa.Right - physW;
                if (newPhysY < wa.Top) newPhysY = wa.Top;
                if (newPhysY + physH > wa.Bottom) newPhysY = wa.Bottom - physH;

                // 3. 如果物理位置有變，計算位移量並套用到邏輯座標
                if (Math.Abs(newPhysX - physX) > 0.5 || Math.Abs(newPhysY - physY) > 0.5)
                {
                    // 取得目前的 DPI 縮放率來將物理位移轉回邏輯位移
                    // 雖然這在跨螢幕時仍有微小誤差，但因為我們是「增量修正」，誤差會被抵消
                    double dx = newPhysX - physX;
                    double dy = newPhysY - physY;

                    // 透過 PointFromScreen 獲取修正後的邏輯座標
                    var targetPhysPos = new System.Windows.Point(newPhysX, newPhysY);
                    var newLogicPosInsideWindow = this.PointFromScreen(targetPhysPos);
                    
                    // 更新 TargetLeft/Top
                    TargetLeft = this.Left + newLogicPosInsideWindow.X;
                    TargetTop = this.Top + newLogicPosInsideWindow.Y;

                    // 同步物理引擎狀態，防止大幅震盪
                    _lastWindowLeft = TargetLeft;
                    _lastWindowTop = TargetTop;
                    _accumulatedDeltaX = 0;
                    _accumulatedDeltaY = 0;

                    foreach (var row in KeyRows)
                    {
                        if (row == null) continue;
                        foreach (var k in row) k?.RandomizePhysics();
                    }
                }
            }
            finally
            {
                _isClamping = false;
            }
        }


        private void CenterWindow()
        {
            var screen = System.Windows.Forms.Screen.PrimaryScreen;
            if (screen == null) return;
            var wa = screen.WorkingArea;

            var source = PresentationSource.FromVisual(this);
            Matrix transform = Matrix.Identity;
            if (source?.CompositionTarget != null) transform = source.CompositionTarget.TransformFromDevice;

            var topLeft = transform.Transform(new System.Windows.Point(wa.Left, wa.Top));
            var bottomRight = transform.Transform(new System.Windows.Point(wa.Right, wa.Bottom));
            double waLeft = topLeft.X;
            double waTop = topLeft.Y;
            double waWidth = Math.Max(0, bottomRight.X - topLeft.X);
            double waHeight = Math.Max(0, bottomRight.Y - topLeft.Y);

            double w = TargetWidth;
            double h = TargetHeight;

            TargetLeft = waLeft + (waWidth - w) / 2.0;
            TargetTop = waTop + (waHeight - h);

            ClampToScreen();
        }

        private void ReduceMemoryUsage()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    using (var proc = System.Diagnostics.Process.GetCurrentProcess())
                    {
                        SetProcessWorkingSetSize(proc.Handle, -1, -1);
                    }
                }
            }
            catch { }
        }

        private void SetupVirtualScreen()
        {
            // 使用 Win32 直接獲取虛擬螢幕邊界，排除 SystemParameters 的快取延遲與偏差
            int physLeft = GetSystemMetrics(76); // SM_XVIRTUALSCREEN
            int physTop = GetSystemMetrics(77);  // SM_YVIRTUALSCREEN
            int physWidth = GetSystemMetrics(78); // SM_CXVIRTUALSCREEN
            int physHeight = GetSystemMetrics(79); // SM_CYVIRTUALSCREEN

            // 確保 MainWindow 覆蓋物理全景，無論熱插拔或解析度如何
            var source = PresentationSource.FromVisual(this);
            double dpiX = 96.0, dpiY = 96.0;
            if (source?.CompositionTarget != null)
            {
                dpiX = 96.0 * source.CompositionTarget.TransformToDevice.M11;
                dpiY = 96.0 * source.CompositionTarget.TransformToDevice.M22;
            }

            this.Left = physLeft * 96.0 / dpiX;
            this.Top = physTop * 96.0 / dpiY;
            this.Width = physWidth * 96.0 / dpiX;
            this.Height = physHeight * 96.0 / dpiY;
        }

        private void MainWindow_Loaded(object? sender, RoutedEventArgs e)
        {
            SetupVirtualScreen();
            
            if (_needsInitialCentering)
            {
                CenterWindow();
                _needsInitialCentering = false;
            }

            // Initial positioning handle
            ClampToScreen();

            // Release memory after initialization
            Task.Delay(1000).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ReduceMemoryUsage());
            });
        }

        private void VisualSyncTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                bool physShift = (GetAsyncKeyState(0x10) & 0x8000) != 0;
                bool physCtrl = (GetAsyncKeyState(0x11) & 0x8000) != 0;
                bool physAlt = (GetAsyncKeyState(0x12) & 0x8000) != 0;
                bool physWin = (GetAsyncKeyState(0x5B) & 0x8000) != 0;

                if (physShift)
                {
                    if (!_lastPhysicalShiftDown) _physShiftDownTime = DateTime.Now;
                }
                else _physShiftDownTime = DateTime.MinValue;

                _lastPhysicalShiftDown = physShift;
                bool isPhysShiftLongPressed = physShift && (DateTime.Now - _physShiftDownTime).TotalMilliseconds > 200;
                
                bool newCapsLock = (GetKeyState(0x14) & 0x0001) != 0;
                bool newNumLock = (GetKeyState(0x90) & 0x0001) != 0;
                bool newShift = isPhysShiftLongPressed || _virtualShiftToggle;
                bool newCtrl = physCtrl || _virtualCtrlToggle;
                bool newAlt = physAlt || _virtualAltToggle;
                bool newWin = physWin || _virtualWinToggle;

                // [重要] 如果狀態發生變更，主動觸發 UpdateDisplay 以確保 UI 反饋即時
                if (newCapsLock != _isCapsLockActive || newNumLock != _isNumLockActive || 
                    newShift != _isShiftActive || newCtrl != _isCtrlActive || 
                    newAlt != _isAltActive || newWin != _isWinActive)
                {
                    _isCapsLockActive = newCapsLock;
                    _isNumLockActive = newNumLock;
                    _isShiftActive = newShift;
                    _isCtrlActive = newCtrl;
                    _isAltActive = newAlt;
                    _isWinActive = newWin;
                    UpdateDisplay(); // 狀態改變，立即重繪
                }

                // --- 視窗隨動與物理演算邏輯 (包含安全性檢查) ---
                double curLeft = TargetLeft;
                double curTop = TargetTop;

                // 若視窗位置尚未初始化或為無效值，則跳過隨動計算
                if (double.IsNaN(curLeft) || double.IsNaN(curTop)) return;

                if (double.IsNaN(_lastWindowLeft))
                {
                    _lastWindowLeft = curLeft;
                    _lastWindowTop = curTop;
                }

                // 累積位移，捕捉極微小移動防止精度丟失
                _accumulatedDeltaX += curLeft - _lastWindowLeft;
                _accumulatedDeltaY += curTop - _lastWindowTop;
                _lastWindowLeft = curLeft;
                _lastWindowTop = curTop;

                double deltaX = 0;
                double deltaY = 0;
                bool hasMoved = false;

                // 提高累積閾值至 0.1，進一步屏蔽 WPF 的座標更新微噪訊
                if (Math.Abs(_accumulatedDeltaX) > 0.1 || Math.Abs(_accumulatedDeltaY) > 0.1)
                {
                    deltaX = _accumulatedDeltaX;
                    deltaY = _accumulatedDeltaY;
                    _accumulatedDeltaX = 0;
                    _accumulatedDeltaY = 0;
                    hasMoved = true;
                }

                // 拖動會話偵測：當剛開始移動時(從靜止變為移動)，重置所有按鍵參數
                if (hasMoved && !_isCurrentlyDragging)
                {
                    foreach (var row in KeyRows)
                    {
                        if (row == null) continue;
                        foreach (var k in row) k?.RandomizePhysics();
                    }
                }
                _isCurrentlyDragging = hasMoved;

                double elapsedSec = (DateTime.Now - _lastFrameTime).TotalSeconds;
                _lastFrameTime = DateTime.Now;
                
                // 絕對安全性：限制每一幀的位移量與時間增量
                if (elapsedSec > 0.1) elapsedSec = 0.03;
                if (elapsedSec < 0.001) elapsedSec = 0.001; 
                double timeFactor = elapsedSec / 0.03;

                // 限制單幀最大位移，但把截斷剩餘的量回存累積器，避免大幅移動時動能瞬間消失造成停頓
                double maxDelta = 150;
                if (_isDragFxEnabled)
                {
                    double clampedX = Math.Clamp(deltaX, -maxDelta, maxDelta);
                    double clampedY = Math.Clamp(deltaY, -maxDelta, maxDelta);
                    // 把未被本幀消化的位移回存，讓下一幀繼續處理
                    _accumulatedDeltaX += deltaX - clampedX;
                    _accumulatedDeltaY += deltaY - clampedY;
                    deltaX = clampedX;
                    deltaY = clampedY;
                }
                else
                {
                    deltaX = 0;
                    deltaY = 0;
                }

                // 物理演算
                bool anyKeyMoving = hasMoved; // 如果視窗還在動，物理一定在動
                foreach (var row in KeyRows)
                {
                    if (row == null) continue;
                    foreach (var k in row)
                    {
                        if (k == null) continue;

                        double offX = k.OffsetX - deltaX;
                        double offY = k.OffsetY - deltaY;

                        // 防止數值異常 (NaN 與 Infinity 防護)
                        if (!double.IsFinite(offX)) offX = 0;
                        if (!double.IsFinite(offY)) offY = 0;
                        
                        double velX = k.VelocityX;
                        double velY = k.VelocityY;
                        if (!double.IsFinite(velX)) velX = 0;
                        if (!double.IsFinite(velY)) velY = 0;

                        // 追隨感核心：改進的彈簧阻尼模型
                        double stiffness = k.FollowSpeed * 0.5; 
                        
                        double accelX = (-offX * stiffness) / k.Mass;
                        double accelY = (-offY * stiffness) / k.Mass;
                        
                        velX += accelX * timeFactor;
                        velY += accelY * timeFactor;

                        velX *= k.Friction;
                        velY *= k.Friction;

                        offX += velX * timeFactor;
                        offY += velY * timeFactor;

                        // 更加激進的歸零機制：提高到 0.5 像素，讓按鍵在肉眼難辨的微動階段就直接歸位
                        if (!hasMoved)
                        {
                            if (Math.Abs(offX) < 0.5 && Math.Abs(velX) < 0.5) { offX = 0; velX = 0; }
                            if (Math.Abs(offY) < 0.5 && Math.Abs(velY) < 0.5) { offY = 0; velY = 0; }
                        }

                        // 提高「忙碌」判定門檻到 1.0 像素。
                        // 這意味著按鍵只要進入最後的細微回彈（小於 1 像素）即視為「靜止」，
                        // 以便立即讓左上角資訊欄恢復，不再等待完全靜止。
                        if (Math.Abs(offX) > 1.0 || Math.Abs(offY) > 1.0 || Math.Abs(velX) > 1.0 || Math.Abs(velY) > 1.0)
                        {
                            anyKeyMoving = true;
                        }

                        // 批量更新，減少 PropertyChanged 觸發次數，防止 WPF 綁定崩潰
                        if (Math.Abs(k.VelocityX - velX) > 0.0001) k.VelocityX = velX;
                        if (Math.Abs(k.VelocityY - velY) > 0.0001) k.VelocityY = velY;
                        if (Math.Abs(k.OffsetX - offX) > 0.001) k.OffsetX = offX;
                        if (Math.Abs(k.OffsetY - offY) > 0.001) k.OffsetY = offY;

                        // 2. 實體按鍵狀態同步
                        if (k.VkCode != MODE_KEY_CODE && k.VkCode != FN_KEY_CODE)
                        {
                            bool isPressed = (GetAsyncKeyState(k.VkCode) & 0x8000) != 0;
                            if (k.VkCode == 0x10) isPressed = physShift;
                            else if (k.VkCode == 0x11) isPressed = physCtrl;
                            else if (k.VkCode == 0x12) isPressed = physAlt;

                            if (k.IsPressed != isPressed) k.IsPressed = isPressed;
                        }
                    }
                }
                _isPhysicsActive = anyKeyMoving;

                // [視覺回饋] 當系統進入抑制狀態時，清空資訊欄方便觀察
                bool currentBusy = IsBusySuppressing;
                if (currentBusy)
                {
                    if (!string.IsNullOrEmpty(CurrentAppInfo))
                    {
                        CurrentAppInfo = "";
                        _lastLoggedAppInfo = ""; // 重置日誌緩存以確保歸位後能重新觸發更新
                    }
                }
                else if (_wasBusySuppressing)
                {
                    // [即時恢復] 當狀態從 Busy 轉為 Idle 的瞬間，立即主動觸發一次偵測，消除 300ms 延遲
                    _lastLoggedAppInfo = ""; 
                    _lastClipboardText = ""; // 強制重置剪貼簿快取，確保恢復後立即抓取最新內容
                    _clipboardCheckCounter = 0;
                    DetectImeStatus();
                }
                _wasBusySuppressing = currentBusy;

                // 僅在狀態真正有變時才更新顯示 (優化效能，防止 UI 管道壅塞並消除字串記憶體配置)
                int currentStateHash = 
                    (_isShiftActive ? 1 : 0) |
                    (_isCapsLockActive ? 2 : 0) |
                    (_isCtrlActive ? 4 : 0) |
                    (_isAltActive ? 8 : 0) |
                    (_isWinActive ? 16 : 0) |
                    (_isZhuyinMode ? 32 : 0) |
                    (_virtualFnToggle ? 64 : 0) |
                    (_temporaryEnglishMode ? 128 : 0);

                if (currentStateHash != _lastDisplayStateHash)
                {
                    _lastDisplayStateHash = currentStateHash;
                    UpdateDisplay();
                }
            }
            catch (Exception ex)
            {
                // 記錄或忽略異常以防止閃退
                System.Diagnostics.Debug.WriteLine($"VisualSyncTimer_Tick Error: {ex.Message}");
            }
        }

        private void DetectImeStatus()
        {
            // [優化] 當鍵盤正在物理運動、任何形式的拖曳或縮放中，暫停實質偵測
            if (IsBusySuppressing) return;

            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero) return;

            uint threadId = GetWindowThreadProcessId(foregroundHwnd, out uint pid);
            string processName = "Unknown";
            
            if (pid == _lastPid && _lastProcessName != null)
            {
                processName = _lastProcessName;
            }
            else
            {
                try
                {
                    using (var proc = System.Diagnostics.Process.GetProcessById((int)pid))
                    {
                        processName = proc.ProcessName;
                    }
                    _lastPid = pid;
                    _lastProcessName = processName;
                }
                catch { }
            }

            GUITHREADINFO guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));

            IntPtr targetHwnd = foregroundHwnd;
            if (GetGUIThreadInfo(threadId, ref guiInfo))
            {
                if (guiInfo.hwndFocus != IntPtr.Zero) targetHwnd = guiInfo.hwndFocus;
            }

            IntPtr imeWnd = ImmGetDefaultIMEWnd(targetHwnd);
            IntPtr hIMC = ImmGetContext(imeWnd != IntPtr.Zero ? imeWnd : targetHwnd);

            bool isChineseMode = _isZhuyinMode;

            if (DateTime.UtcNow < _ignoreImeSyncUntil)
            {
                if (hIMC != IntPtr.Zero) ImmReleaseContext(targetHwnd, hIMC);
                UpdateAppInfo(processName, isChineseMode);
                return;
            }

            bool isChinese = false;
            bool statusRead = false;

            // 優先使用 WM_IME_CONTROL 直接詢問 IME 視窗 (能支援 Weasel 小狼毫等 TSF，此時 hIMC 可能為 Zero)
            if (imeWnd != IntPtr.Zero)
            {
                IntPtr resOpen = SendMessage(imeWnd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero);
                IntPtr resConv = SendMessage(imeWnd, WM_IME_CONTROL, (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero);

                bool isOpen = resOpen.ToInt32() != 0;
                bool isNative = (resConv.ToInt32() & IME_CMODE_NATIVE) != 0;

                isChinese = isOpen && isNative;
                statusRead = true;
            }
            // 若沒有 imeWnd 但有傳統 hIMC，則退回使用舊版 API (例如舊應用)
            else if (hIMC != IntPtr.Zero)
            {
                bool isOpen = ImmGetOpenStatus(hIMC);
                if (ImmGetConversionStatus(hIMC, out uint conv, out uint sentence))
                {
                    bool isNative = (conv & IME_CMODE_NATIVE) != 0;
                    isChinese = isOpen && isNative;
                    statusRead = true;
                }
            }

            if (statusRead)
            {
                isChineseMode = isChinese;
                if (_isZhuyinMode != isChinese)
                {
                    _isZhuyinMode = isChinese;
                    _localPreviewToggle = false;
                    _temporaryEnglishMode = false;
                    UpdateDisplay();
                }
            }

            if (hIMC != IntPtr.Zero)
            {
                ImmReleaseContext(targetHwnd, hIMC);
            }

            UpdateAppInfo(processName, isChineseMode);
        }

        private string GetClipboardTextSafe()
        {
            try
            {
                if (System.Windows.Clipboard.ContainsImage())
                {
                    return "🖼️ 圖片";
                }
                else if (System.Windows.Clipboard.ContainsFileDropList())
                {
                    var files = System.Windows.Clipboard.GetFileDropList();
                    if (files != null && files.Count > 0)
                    {
                        string fileName = System.IO.Path.GetFileName(files[0]) ?? "";
                        if (files.Count > 1)
                        {
                            return $"📁 {fileName} 等 {files.Count} 個檔案";
                        }
                        return $"📁 {fileName}";
                    }
                }
                else if (System.Windows.Clipboard.ContainsText())
                {
                    string text = System.Windows.Clipboard.GetText();
                    if (!string.IsNullOrEmpty(text))
                    {
                        text = text.Replace("\r", "").Replace("\n", " ").Replace("\t", " ").Trim();
                        if (text.Length > 10)
                        {
                            text = text.Substring(0, 10) + "...";
                        }
                        return text;
                    }
                }
            }
            catch
            {
                // Ignore clipboard access exceptions
            }
            return "";
        }

        private string GetClipboardTextSafelyThrottled()
        {
            _clipboardCheckCounter++;
            if (_clipboardCheckCounter >= 10 || string.IsNullOrEmpty(_lastClipboardText))
            {
                _clipboardCheckCounter = 0;
                _lastClipboardText = GetClipboardTextSafe();
            }
            return _lastClipboardText;
        }

        private void UpdateAppInfo(string processName, bool isZhuyin)
        {
            string status = isZhuyin ? "注音" : "英文";
            string clip = GetClipboardTextSafelyThrottled();
            string newInfo = $"{processName} - {status}";
            if (!string.IsNullOrEmpty(clip))
            {
                newInfo += $" | 📋 {clip}";
            }

            if (_lastLoggedAppInfo != newInfo)
            {
                _lastLoggedAppInfo = newInfo;
                // [優化] 直接設定屬性，避免 Dispatcher 造成的非同步微小延遲
                CurrentAppInfo = newInfo;
            }
        }

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var helper = new WindowInteropHelper(this);
            SetWindowLong(helper.Handle, GWL_EXSTYLE, GetWindowLong(helper.Handle, GWL_EXSTYLE) | WS_EX_NOACTIVATE);

            // 強制關閉 Windows 11 預設的視窗轉場效果 (例如觸控時的 Cloud White Shrink / Snap 預覽)
            int disableTransitions = 1;
            DwmSetWindowAttribute(helper.Handle, DWMWA_TRANSITIONS_FORCEDISABLED, ref disableTransitions, sizeof(int));

            // 嘗試保留圓角 (如果系統支援)
            int cornerPref = DWMWCP_ROUND;
            DwmSetWindowAttribute(helper.Handle, DWMWA_WINDOW_CORNER_PREFERENCE, ref cornerPref, sizeof(int));

            HwndSource.FromHwnd(helper.Handle).AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_DISPLAYCHANGE = 0x007E;
            const int WM_DPICHANGED = 0x02E0;
            const int WM_SETTINGCHANGE = 0x001A;

            if (msg == WM_DISPLAYCHANGE || msg == WM_DPICHANGED)
            {
                // [修正] 雙重保險：監聽螢幕變更與 DPI 變更，即時更新虛擬畫布
                SetupVirtualScreen();
                ClampToScreen();
                handled = true;
            }

            if (msg == WM_SETTINGCHANGE)
            {
                // [新功能] 當系統偏好設定（如深淺色模式）變更時，自動偵測並同步主題
                DetectSystemTheme();
            }
            if (msg == WM_NCHITTEST)
            {
                if (_isTouchDraggingWindow)
                {
                    handled = true;
                    return (IntPtr)HTCLIENT;
                }
            }

            if (msg == _msgShowOsk) { ToggleVisibility(); handled = true; }
            if (msg == WM_WINDOWPOSCHANGING)
            {
                // [修正] 移除主視窗邊界限制邏輯。
                // 此主視窗 (MainWindow) 作為透明底層，必須覆蓋整個虛擬桌面區域 (SystemParameters.VirtualScreen)。
                // 若在此處將其限制在單一螢幕的 WorkingArea 內，鍵盤將無法跨越螢幕邊界，進而造成「三螢幕回彈」問題。
            }
            return IntPtr.Zero;
        }

        // Win32 API for native Monitor matching
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromRect([In] ref RECT lprc, uint dwFlags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int left; public int top; public int right; public int bottom; }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MONITORINFOEX
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public int dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szDevice;
        }

        private void ToggleImeStatus(bool targetIsChinese)
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero)
            {
                FallbackImeToggle();
                return;
            }

            uint threadId = GetWindowThreadProcessId(foregroundHwnd, out _);
            GUITHREADINFO guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));

            IntPtr targetHwnd = foregroundHwnd;
            if (GetGUIThreadInfo(threadId, ref guiInfo))
            {
                if (guiInfo.hwndFocus != IntPtr.Zero) targetHwnd = guiInfo.hwndFocus;
            }

            IntPtr imeWnd = ImmGetDefaultIMEWnd(targetHwnd);
            IntPtr hIMC = ImmGetContext(imeWnd != IntPtr.Zero ? imeWnd : targetHwnd);

            if (hIMC != IntPtr.Zero)
            {
                if (ImmGetConversionStatus(hIMC, out uint conv, out uint sentence))
                {
                    if (targetIsChinese) conv |= IME_CMODE_NATIVE;
                    else conv &= ~IME_CMODE_NATIVE;

                    ImmSetConversionStatus(hIMC, conv, sentence);
                }
                ImmReleaseContext(targetHwnd, hIMC);
            }
            else
            {
                FallbackImeToggle();
            }
        }

        private void FallbackImeToggle()
        {
            var inputs = new List<INPUT>();
            // Fallback utilizing proper ScanCodes so the IME respects it
            var inputDown = new INPUT { type = INPUT_KEYBOARD, u = new InputUnion { ki = new KEYBDINPUT { wVk = 0xA0, wScan = 0x2A, dwFlags = 0x0008, time = 0, dwExtraInfo = IntPtr.Zero } } };
            var inputUp = new INPUT { type = INPUT_KEYBOARD, u = new InputUnion { ki = new KEYBDINPUT { wVk = 0xA0, wScan = 0x2A, dwFlags = 0x0008 | KEYEVENTF_KEYUP, time = 0, dwExtraInfo = IntPtr.Zero } } };
            inputs.Add(inputDown);
            inputs.Add(inputUp);
            SendInput((uint)inputs.Count, inputs.ToArray(), INPUT.Size);
        }

        private void ToggleVisibility()
        {
            if (this.Visibility == Visibility.Visible)
            {
                this.Hide();
                Task.Delay(500).ContinueWith(_ => this.Dispatcher.Invoke(() => ReduceMemoryUsage()));
            }
            else { this.Show(); }
        }

        private void OnKeyClick(KeyModel? key)
        {
            if (key == null || IsEditMode) return; // 編輯模式下停用快速點擊功能

            if (key.VkCode == MODE_KEY_CODE)
            {
                if (_modeKeyLongPressHandled) { _modeKeyLongPressHandled = false; return; }

                if (_virtualFnToggle)
                {
                    _localPreviewToggle = !_localPreviewToggle;
                    _temporaryEnglishMode = false;
                    _virtualShiftToggle = false;
                    UpdateDisplay();
                    return;
                }

                ToggleImeStatus(!_isZhuyinMode);
                _isZhuyinMode = !_isZhuyinMode;
                _ignoreImeSyncUntil = DateTime.UtcNow.AddMilliseconds(300);
                _temporaryEnglishMode = false;
                ResetVirtualModifiers();
                UpdateDisplay();
                return;
            }

            if (key.VkCode == CAD_KEY_CODE)
            {
                ShowSecurityMenu();
                if (!_temporaryEnglishMode) ResetVirtualModifiers();
                UpdateDisplay();
                return;
            }

            if (key.VkCode == FN_KEY_CODE) { _virtualFnToggle = !_virtualFnToggle; UpdateDisplay(); return; }
            if (key.VkCode == 0x11 || key.VkCode == 0xA2 || key.VkCode == 0xA3) { _virtualCtrlToggle = !_virtualCtrlToggle; if (_virtualCtrlToggle) _activeCtrlVk = key.VkCode; UpdateDisplay(); return; }
            if (key.VkCode == 0x12 || key.VkCode == 0xA4 || key.VkCode == 0xA5)
            {
                _virtualAltToggle = !_virtualAltToggle;
                if (_virtualAltToggle) { _activeAltVk = key.VkCode; SendSimulatedKey(key.VkCode, false); } else SendSimulatedKey(_activeAltVk, true);
                UpdateDisplay();
                return;
            }
            if (key.VkCode == 0x5B || key.VkCode == 0x5C) { _virtualWinToggle = !_virtualWinToggle; if (_virtualWinToggle) _activeWinVk = key.VkCode; UpdateDisplay(); return; }

            if ((key.VkCode == 0x10 || key.VkCode == 0xA0 || key.VkCode == 0xA1) && _isZhuyinMode)
            {
                _temporaryEnglishMode = !_temporaryEnglishMode;
                UpdateDisplay();
                return;
            }
            if (key.VkCode == 0x10 || key.VkCode == 0xA0 || key.VkCode == 0xA1) { _virtualShiftToggle = !_virtualShiftToggle; if (_virtualShiftToggle) _activeShiftVk = key.VkCode; UpdateDisplay(); return; }

            if (_virtualFnToggle && key.VkCode == 0x09)
            {
                ShowSecurityMenu();
                if (!_temporaryEnglishMode) ResetVirtualModifiers();
                UpdateDisplay();
                return;
            }

            byte sendVk = key.VkCode;
            if (_virtualFnToggle && FnSendMap.TryGetValue(key.VkCode, out byte targetVk)) sendVk = targetVk;
            if (_virtualAltToggle && key.VkCode >= 0x30 && key.VkCode <= 0x39) sendVk = (byte)(0x60 + (key.VkCode - 0x30));

            if (sendVk == 0x2E && _isCtrlActive && _isAltActive)
            {
                TryStartTaskManager();
                if (!_temporaryEnglishMode) ResetVirtualModifiers();
                UpdateDisplay();
                return;
            }

            bool physShift = (GetAsyncKeyState(0x10) & 0x8000) != 0;
            bool needInjectShift = _virtualShiftToggle;

            if (_temporaryEnglishMode)
            {
                needInjectShift = true;
            }

            bool effectiveShiftInject = needInjectShift && !physShift;

            bool isNumpadKey = (sendVk >= 0x60 && sendVk <= 0x69 || sendVk == 0x6E);
            if (isNumpadKey && effectiveShiftInject)
            {
                // Do not inject virtual shift for Numpad keys, so it doesn't artificially reverse NumLock logic
                effectiveShiftInject = false;
            }

            var inputs = new List<INPUT>();

            if (_virtualCtrlToggle) AddKeyInput(inputs, _activeCtrlVk, false);
            if (_virtualWinToggle) AddKeyInput(inputs, _activeWinVk, false);

            if (effectiveShiftInject) AddKeyInput(inputs, _activeShiftVk, false);

            AddKeyInput(inputs, sendVk, false);
            AddKeyInput(inputs, sendVk, true);

            if (effectiveShiftInject) AddKeyInput(inputs, _activeShiftVk, true);

            if (_virtualWinToggle) AddKeyInput(inputs, _activeWinVk, true);
            if (_virtualCtrlToggle) AddKeyInput(inputs, _activeCtrlVk, true);

            if (inputs.Count > 0)
            {
                SendInput((uint)inputs.Count, inputs.ToArray(), INPUT.Size);
            }

            if (sendVk != 0x11 && sendVk != 0xA2 && sendVk != 0xA3 &&
                sendVk != 0x12 && sendVk != 0xA4 && sendVk != 0xA5 &&
                sendVk != 0x10 && sendVk != 0xA0 && sendVk != 0xA1 &&
                sendVk != 0x5B && sendVk != 0x5C)
            {
                if (!_temporaryEnglishMode)
                {
                    ResetVirtualModifiers();
                }
            }

            if (_temporaryEnglishMode && key.VkCode == 0x0D)
            {
                _temporaryEnglishMode = false;
            }

            UpdateDisplay();
        }

        private void ApplyJellyAnimation(UIElement element, double toScaleX, double toScaleY, double durationMs, bool elastic)
        {
            ScaleTransform? transform = null;
            if (element.RenderTransform is TransformGroup group)
            {
                transform = group.Children.OfType<ScaleTransform>().FirstOrDefault();
            }
            if (transform == null)
            {
                transform = element.RenderTransform as ScaleTransform;
                if (transform == null || element.RenderTransformOrigin.X != 0.5)
                {
                    transform = new ScaleTransform(1, 1);
                    element.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
                    element.RenderTransform = transform;
                }
            }

            transform.BeginAnimation(ScaleTransform.ScaleXProperty, null);
            transform.BeginAnimation(ScaleTransform.ScaleYProperty, null);

            var animX = new DoubleAnimation { To = toScaleX, Duration = TimeSpan.FromMilliseconds(durationMs) };
            var animY = new DoubleAnimation { To = toScaleY, Duration = TimeSpan.FromMilliseconds(durationMs) };

            if (elastic)
            {
                // 超級 Q 彈的果凍回彈 (彈簧更軟、震盪更多次)
                animX.EasingFunction = new ElasticEase { EasingMode = EasingMode.EaseOut, Oscillations = 4, Springiness = 3 };
                animY.EasingFunction = new ElasticEase { EasingMode = EasingMode.EaseOut, Oscillations = 4, Springiness = 3 };
            }
            else
            {
                // 瞬間壓扁的緩動
                animX.EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut };
                animY.EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut };
            }

            transform.BeginAnimation(ScaleTransform.ScaleXProperty, animX);
            transform.BeginAnimation(ScaleTransform.ScaleYProperty, animY);
        }

        private void KeyBorder_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (sender is System.Windows.Controls.Border b)
            {
                // Give it a subtle bump when hovering over
                ApplyJellyAnimation(b, 1.05, 1.05, 500, true);
            }
        }

        private void KeyBorder_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return; // 忽略被 WPF 轉換的觸控假滑鼠事件
            if (IsEditMode) return; // 進入編輯模式時停用任何觸發特效與按鍵行為

            if (sender is System.Windows.Controls.Border b)
            {
                b.Opacity = 0.8;
                // 真·果凍擠壓：寬度變寬，高度變扁 (體積守恆錯覺)
                ApplyJellyAnimation(b, 1.12, 0.85, 100, false);
            }
            if ((sender as FrameworkElement)?.DataContext is KeyModel key) { e.Handled = true; OnKeyClick(key); }
        }

        private void KeyBorder_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (sender is System.Windows.Controls.Border b)
            {
                b.ClearValue(UIElement.OpacityProperty);
                // 誇張的果凍回彈
                ApplyJellyAnimation(b, 1.0, 1.0, 1000, true);
            }
        }

        private void KeyBorder_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (sender is System.Windows.Controls.Border b)
            {
                b.ClearValue(UIElement.OpacityProperty);
                // Return to normal size cleanly
                ApplyJellyAnimation(b, 1.0, 1.0, 300, false);
            }
        }

        private void KeyBorder_TouchDown(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (IsEditMode) return; // 進入編輯模式時停用所有輸入交互
            if (sender is System.Windows.Controls.Border b)
            {
                b.Opacity = 0.8;
                // 真·果凍擠壓
                ApplyJellyAnimation(b, 1.12, 0.85, 100, false);
            }
            if ((sender as FrameworkElement)?.DataContext is KeyModel key) { e.Handled = true; OnKeyClick(key); }
        }

        private void KeyBorder_TouchUp(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (sender is System.Windows.Controls.Border b)
            {
                b.ClearValue(UIElement.OpacityProperty);
                ApplyJellyAnimation(b, 1.0, 1.0, 1000, true);
            }
        }

        private void KeyBorder_TouchLeave(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (sender is System.Windows.Controls.Border b)
            {
                b.ClearValue(UIElement.OpacityProperty);
                ApplyJellyAnimation(b, 1.0, 1.0, 300, false);
            }
        }

        #region 拖曳交換按鈕邏輯
        private void KeyBorder_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!IsEditMode) return;
            if (e.StylusDevice != null) return;
            if (sender is System.Windows.Controls.Border border && border.DataContext is KeyModel key)
            {
                e.Handled = true;
                _dragStartPoint = e.GetPosition(this);
                _draggedBorder = border;
                _draggedKey = key;
                _isManualDragging = false;
                border.CaptureMouse();
            }
        }

        private void KeyBorder_PreviewTouchDown(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (!IsEditMode) return;
            if (sender is System.Windows.Controls.Border border && border.DataContext is KeyModel key)
            {
                e.Handled = true;
                _dragStartPoint = e.GetTouchPoint(this).Position;
                _draggedBorder = border;
                _draggedKey = key;
                _isManualDragging = false;
                e.TouchDevice.Capture(border);
            }
        }

        private void Window_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isMouseDraggingWindow)
            {
                var pos = e.GetPosition(this);
                double dx = pos.X - _mouseDragStartPoint.X;
                double dy = pos.Y - _mouseDragStartPoint.Y;

                TargetLeft += dx;
                TargetTop += dy;
                _mouseDragStartPoint = pos;
            }

            if (!IsEditMode || _draggedKey == null || _draggedBorder == null) return;
            if (e.StylusDevice != null) return;

            if (e.LeftButton == MouseButtonState.Pressed)
            {
                HandleDragMove(e.GetPosition(this), e.GetPosition(DragOverlay));
            }
        }

        private void Window_PreviewTouchMove(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (!IsEditMode || _draggedKey == null || _draggedBorder == null) return;
            HandleDragMove(e.GetTouchPoint(this).Position, e.GetTouchPoint(DragOverlay).Position);
        }

        private void HandleDragMove(System.Windows.Point posInWindow, System.Windows.Point posInOverlay)
        {
            if (!_isManualDragging)
            {
                Vector diff = _dragStartPoint - posInWindow;
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    StartManualDrag(posInOverlay);
                }
            }

            if (_isManualDragging && _dragGhost != null)
            {
                Canvas.SetLeft(_dragGhost, posInOverlay.X - _ghostOffset.X);
                Canvas.SetTop(_dragGhost, posInOverlay.Y - _ghostOffset.Y);
            }
        }

        private void StartManualDrag(System.Windows.Point posInOverlay)
        {
            if (_draggedBorder == null) return;
            _isManualDragging = true;

            var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap((int)Math.Max(1, _draggedBorder.ActualWidth), (int)Math.Max(1, _draggedBorder.ActualHeight), 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(_draggedBorder);
            var brush = new System.Windows.Media.ImageBrush(rtb) { Stretch = System.Windows.Media.Stretch.None };

            _dragGhost = new System.Windows.Shapes.Rectangle
            {
                Width = _draggedBorder.ActualWidth,
                Height = _draggedBorder.ActualHeight,
                Fill = brush,
                Opacity = 0.8
            };

            System.Windows.Point borderPos = _draggedBorder.TranslatePoint(new System.Windows.Point(0, 0), DragOverlay);
            _ghostOffset = new System.Windows.Point(posInOverlay.X - borderPos.X, posInOverlay.Y - borderPos.Y);

            Canvas.SetLeft(_dragGhost, borderPos.X);
            Canvas.SetTop(_dragGhost, borderPos.Y);

            var rTrans = new RotateTransform();
            _dragGhost.RenderTransform = rTrans;
            _dragGhost.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
            var anim = new DoubleAnimation(-4, 4, new Duration(TimeSpan.FromMilliseconds(50))) { AutoReverse = true, RepeatBehavior = RepeatBehavior.Forever };
            rTrans.BeginAnimation(RotateTransform.AngleProperty, anim);

            DragOverlay.Children.Add(_dragGhost);
            _draggedBorder.Opacity = 0.0;
        }

        private void Window_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isMouseDraggingWindow)
            {
                _isMouseDraggingWindow = false;
                ClampToScreen();
                Mouse.Capture(null);
                SaveSettings(); // 拖曳結束後存檔
            }

            if (!IsEditMode) return;
            if (e.StylusDevice != null) return;
            HandleDrop(e.GetPosition(MainRootGrid));
        }

        private void Window_PreviewTouchUp(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (!IsEditMode) return;
            HandleDrop(e.GetTouchPoint(MainRootGrid).Position);
        }

        private async void HandleDrop(System.Windows.Point dropPt)
        {
            if (_isManualDragging && _draggedBorder != null && _dragGhost != null && _draggedKey != null)
            {
                _isManualDragging = false;
                _draggedBorder.ReleaseMouseCapture();
                _draggedBorder.ReleaseAllTouchCaptures();

                System.Windows.Controls.Border? targetBorder = null;
                KeyModel? targetKey = null;

                // Perform HitTest without picking up the Ghost rectangle itself
                VisualTreeHelper.HitTest(MainRootGrid, null, new HitTestResultCallback(result =>
                {
                    DependencyObject visualHit = result.VisualHit;
                    while (visualHit != null && !(visualHit is System.Windows.Controls.Border))
                    {
                        visualHit = VisualTreeHelper.GetParent(visualHit);
                    }

                    if (visualHit is System.Windows.Controls.Border b && b.DataContext is KeyModel km && b != _draggedBorder)
                    {
                        targetBorder = b;
                        targetKey = km;
                        return HitTestResultBehavior.Stop;
                    }

                    return HitTestResultBehavior.Continue;
                }), new PointHitTestParameters(dropPt));

                if (targetBorder != null && targetKey != null)
                {
                    _dragGhost.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, null);

                    var targetRtb = new System.Windows.Media.Imaging.RenderTargetBitmap((int)Math.Max(1, targetBorder.ActualWidth), (int)Math.Max(1, targetBorder.ActualHeight), 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                    targetRtb.Render(targetBorder);
                    var targetBrush = new System.Windows.Media.ImageBrush(targetRtb) { Stretch = System.Windows.Media.Stretch.None };
                    var targetGhost = new System.Windows.Shapes.Rectangle
                    {
                        Width = targetBorder.ActualWidth,
                        Height = targetBorder.ActualHeight,
                        Fill = targetBrush
                    };

                    System.Windows.Point targetPos = targetBorder.TranslatePoint(new System.Windows.Point(0, 0), DragOverlay);
                    System.Windows.Point startGhostPos = new System.Windows.Point(Canvas.GetLeft(_dragGhost), Canvas.GetTop(_dragGhost));

                    Canvas.SetLeft(targetGhost, targetPos.X);
                    Canvas.SetTop(targetGhost, targetPos.Y);
                    DragOverlay.Children.Add(targetGhost);

                    targetBorder.Opacity = 0.0;

                    var duration = TimeSpan.FromMilliseconds(200);
                    var ease = new CubicEase { EasingMode = EasingMode.EaseOut };

                    var animX1 = new DoubleAnimation(startGhostPos.X, targetPos.X, duration) { EasingFunction = ease };
                    var animY1 = new DoubleAnimation(startGhostPos.Y, targetPos.Y, duration) { EasingFunction = ease };
                    var animX2 = new DoubleAnimation(targetPos.X, startGhostPos.X, duration) { EasingFunction = ease };
                    var animY2 = new DoubleAnimation(targetPos.Y, startGhostPos.Y, duration) { EasingFunction = ease };

                    _dragGhost.BeginAnimation(Canvas.LeftProperty, animX1);
                    _dragGhost.BeginAnimation(Canvas.TopProperty, animY1);

                    targetGhost.BeginAnimation(Canvas.LeftProperty, animX2);
                    targetGhost.BeginAnimation(Canvas.TopProperty, animY2);

                    await System.Threading.Tasks.Task.Delay(210);

                    SwapKeys(_draggedKey, targetKey);
                    SaveSettings();

                    targetBorder.Opacity = 1.0;
                }
                else
                {
                    _dragGhost.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, null);
                    System.Windows.Point origPos = _draggedBorder.TranslatePoint(new System.Windows.Point(0, 0), DragOverlay);

                    var duration = TimeSpan.FromMilliseconds(200);
                    var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
                    var animX = new DoubleAnimation(Canvas.GetLeft(_dragGhost), origPos.X, duration) { EasingFunction = ease };
                    var animY = new DoubleAnimation(Canvas.GetTop(_dragGhost), origPos.Y, duration) { EasingFunction = ease };

                    _dragGhost.BeginAnimation(Canvas.LeftProperty, animX);
                    _dragGhost.BeginAnimation(Canvas.TopProperty, animY);

                    await System.Threading.Tasks.Task.Delay(210);
                }

                DragOverlay.Children.Clear();
                _draggedBorder.Opacity = 1.0;
                _draggedBorder = null;
                _draggedKey = null;
                _dragGhost = null;
            }
            else
            {
                if (_draggedBorder != null)
                {
                    _draggedBorder.ReleaseMouseCapture();
                    _draggedBorder.ReleaseAllTouchCaptures();
                }
                _draggedBorder = null;
                _draggedKey = null;
            }
        }

        private void SwapKeys(KeyModel k1, KeyModel k2)
        {
            var tempWidth = k1.Width; k1.Width = k2.Width; k2.Width = tempWidth;
            var tempHeight = k1.Height; k1.Height = k2.Height; k2.Height = tempHeight;
            var tempLeftM = k1.LeftMargin; k1.LeftMargin = k2.LeftMargin; k2.LeftMargin = tempLeftM;
            var tempTopM = k1.TopMargin; k1.TopMargin = k2.TopMargin; k2.TopMargin = tempTopM;
            var tempAlign = k1.IsCenterAligned; k1.IsCenterAligned = k2.IsCenterAligned; k2.IsCenterAligned = tempAlign;

            ObservableCollection<KeyModel>? r1 = null;
            int idx1 = -1;
            ObservableCollection<KeyModel>? r2 = null;
            int idx2 = -1;

            foreach (var row in KeyRows)
            {
                if (r1 == null) { int i = row.IndexOf(k1); if (i >= 0) { r1 = row; idx1 = i; } }
                if (r2 == null) { int i = row.IndexOf(k2); if (i >= 0) { r2 = row; idx2 = i; } }
                if (r1 != null && r2 != null) break;
            }

            if (r1 != null && r2 != null)
            {
                if (r1 == r2)
                {
                    r1.Move(idx1, idx2);
                    if (idx1 < idx2) r1.Move(idx2 - 1, idx1);
                    else r1.Move(idx2 + 1, idx1);
                }
                else
                {
                    r1.RemoveAt(idx1);
                    r1.Insert(idx1, k2);
                    r2.RemoveAt(idx2);
                    r2.Insert(idx2, k1);
                }
            }

            UpdateDisplay();
        }
        #endregion

        private void UpdateDisplay()
        {
            bool upper = _isCapsLockActive ^ _isShiftActive;
            bool symbols = _isShiftActive;

            // 真正的中文輸入狀態是：是在注音模式下 && 沒有被 Shift 切換成臨時英文
            bool isActuallyChinese = _isZhuyinMode && !_temporaryEnglishMode;
            bool displayZhuyin = _localPreviewToggle ? !isActuallyChinese : isActuallyChinese;

            ModeIndicator = displayZhuyin ? "En" : "ㄅ";

            if (displayZhuyin)
            {
                IndicatorColor = _isDarkMode ? "Orange" : "#D2691E";
            }
            else
            {
                IndicatorColor = _isDarkMode ? "White" : "#333333";
            }

            if (_temporaryEnglishMode && _isZhuyinMode)
            {
                // 如果是在微軟注音下的英文模式 (Shift 切換)，用特定顏色標註
                IndicatorColor = _themeActiveColor;
            }

            foreach (var row in KeyRows)
            {
                foreach (var k in row)
                {
                    if (k.VkCode == MODE_KEY_CODE)
                    {
                        k.English = ModeIndicator;
                        k.DisplayName = ModeIndicator;
                        k.TextColor = IndicatorColor;
                        k.DynamicTextColor = IndicatorColor;
                        continue;
                    }

                    if (!_isFullLayout && _virtualFnToggle && FnDisplayMap.TryGetValue(k.VkCode, out string? fnLabel))
                    {
                        if (!string.IsNullOrEmpty(fnLabel))
                        {
                            k.DisplayName = fnLabel!;
                            k.DynamicTextColor = _themeSubColor;
                            continue;
                        }
                    }

                    if (k.VkCode == FN_KEY_CODE)
                    {
                        k.TextColor = _isDarkMode ? "#32CD32" : "#2E8B57";
                        k.IsActiveToggle = _virtualFnToggle;
                        k.DisplayName = "⌨";
                        k.DynamicTextColor = _virtualFnToggle ? _themeActiveColor : _themeTextColor;
                        continue;
                    }

                    // 處理注音模式下的 Shift 狀態或臨時英文模式
                    // 為了統一視覺體驗與修飾鍵行為，兩者公用一套符號映射邏輯
                    bool isZhuyinShiftOrTempEnglish = _isZhuyinMode && (_isShiftActive || _temporaryEnglishMode);

                    if (isZhuyinShiftOrTempEnglish)
                    {
                        string? overrideLabel = null;
                        switch (k.VkCode)
                        {
                            case 0xC0: overrideLabel = "～"; break;
                            case 0x31: overrideLabel = "！"; break;
                            case 0x32: overrideLabel = "＠"; break;
                            case 0x33: overrideLabel = "＃"; break;
                            case 0x34: overrideLabel = "＄"; break;
                            case 0x35: overrideLabel = "％"; break;
                            case 0x36: overrideLabel = "^"; break;
                            case 0x37: overrideLabel = "＆"; break;
                            case 0x38: overrideLabel = "＊"; break;
                            case 0x39: overrideLabel = "（"; break;
                            case 0x30: overrideLabel = "）"; break;
                            case 0xBD: overrideLabel = "＿"; break;
                            case 0xBB: overrideLabel = "＋"; break;
                            case 0xDB: overrideLabel = "〔"; break;
                            case 0xDD: overrideLabel = "〕"; break;
                            case 0xDC: overrideLabel = "｜"; break;
                            case 0xDE: overrideLabel = "；"; break;
                            case 0xBA: overrideLabel = "："; break;
                            case 0xBC: overrideLabel = "，"; break;
                            case 0xBE: overrideLabel = "。"; break;
                            case 0xBF: overrideLabel = "？"; break;
                        }

                        if (overrideLabel != null)
                        {
                            k.DisplayName = overrideLabel;
                            k.DynamicTextColor = _themeSubColor;
                        }
                        else
                        {
                            // 字母始終顯示大寫並使用副色 (當在臨時模式或按住 Shift 時)
                            bool isLetter = k.VkCode >= 0x41 && k.VkCode <= 0x5A;
                            k.DisplayName = (isLetter || upper) ? k.EnglishUpper : k.English;
                            k.DynamicTextColor = (isLetter || _temporaryEnglishMode) ? _themeSubColor : _themeTextColor;
                        }
                    }
                    else if (_isZhuyinMode && !string.IsNullOrEmpty(k.Zhuyin))
                    {
                        k.DisplayName = k.Zhuyin;
                        k.DynamicTextColor = _isDarkMode ? "Orange" : "#D2691E";
                    }
                    else
                    {
                        bool isAlpha = k.VkCode >= 0x41 && k.VkCode <= 0x5A;
                        bool isNumpadKey = (k.VkCode >= 0x60 && k.VkCode <= 0x69 || k.VkCode == 0x6E);

                        if (isNumpadKey)
                        {
                            bool showNav = !_isNumLockActive;
                            if (showNav && !string.IsNullOrEmpty(k.EnglishUpper))
                            {
                                k.DisplayName = k.EnglishUpper;
                                // 導航模式 (NumLock OFF)：白字 (一般文字顏色)
                                k.DynamicTextColor = _themeTextColor;
                            }
                            else
                            {
                                k.DisplayName = k.English;
                                // 數字模式 (NumLock ON)：藍字 (與亮起的 Num 鍵同為主題色，解決使用者感受到的版面相反問題)
                                k.DynamicTextColor = _themeNumColor;
                            }
                        }
                        else
                        {
                            k.DisplayName = (isAlpha ? upper : symbols) ? k.EnglishUpper : k.English;

                            bool hasShiftValue = !string.IsNullOrEmpty(k.EnglishUpper) && k.EnglishUpper != k.English;
                            bool shouldHighlightShift = symbols && hasShiftValue;

                            if (isAlpha)
                            {
                                k.DynamicTextColor = upper ? _themeSubColor : _themeTextColor;
                            }
                            else
                            {
                                k.DynamicTextColor = shouldHighlightShift ? _themeSubColor : _themeTextColor;
                            }
                        }
                    }

                    if (k.VkCode == 0x14) { k.TextColor = "Cyan"; k.IsActiveToggle = _isCapsLockActive; k.DisplayName = "⇪"; k.DynamicTextColor = _isCapsLockActive ? "Cyan" : _themeTextColor; }
                    if (k.VkCode == 0x90) { k.TextColor = _themeNumColor; k.IsActiveToggle = _isNumLockActive; k.DisplayName = "Num"; k.DynamicTextColor = _isNumLockActive ? _themeNumColor : _themeTextColor; }
                    if (k.VkCode == 0x10 || k.VkCode == 0xA0 || k.VkCode == 0xA1) { k.TextColor = _isDarkMode ? "#1E90FF" : "#0078D7"; k.IsActiveToggle = _isShiftActive; k.DisplayName = "⇧"; k.DynamicTextColor = _isShiftActive ? (_isDarkMode ? "#1E90FF" : "#0078D7") : _themeTextColor; }
                    if (k.VkCode == 0x11 || k.VkCode == 0xA2 || k.VkCode == 0xA3) { k.TextColor = "Cyan"; k.IsActiveToggle = _isCtrlActive; k.DisplayName = "⌃"; k.DynamicTextColor = _isCtrlActive ? "Cyan" : _themeTextColor; }
                    if (k.VkCode == 0x12 || k.VkCode == 0xA4 || k.VkCode == 0xA5) { k.TextColor = "Cyan"; k.IsActiveToggle = _isAltActive; k.DisplayName = "⌥"; k.DynamicTextColor = _isAltActive ? "Cyan" : _themeTextColor; }
                    if (k.VkCode == 0x5B || k.VkCode == 0x5C) { k.TextColor = "Cyan"; k.IsActiveToggle = _isWinActive; k.DisplayName = "⊞"; k.DynamicTextColor = _isWinActive ? "Cyan" : _themeTextColor; }
                }
            }
        }

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
            r1.Add(new KeyModel { English = "⌫", EnglishUpper = "⌫", Zhuyin = "", VkCode = 0x08, Width = 95, IsCenterAligned = true });
            KeyRows.Add(r1);

            var r2 = new ObservableCollection<KeyModel>();
            r2.Add(new KeyModel { English = "⇥", EnglishUpper = "Tab", Zhuyin = "", VkCode = 0x09, Width = 95, IsCenterAligned = true });
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
            r2.Add(new KeyModel { English = "[", EnglishUpper = "{", Zhuyin = "「", VkCode = 0xDB });
            r2.Add(new KeyModel { English = "]", EnglishUpper = "}", Zhuyin = "」", VkCode = 0xDD });
            r2.Add(new KeyModel { English = "\\", EnglishUpper = "|", Zhuyin = "、", VkCode = 0xDC, Width = 65 });
            KeyRows.Add(r2);

            var r3 = new ObservableCollection<KeyModel>();
            r3.Add(new KeyModel { English = "⇪", EnglishUpper = "Caps", Zhuyin = "", VkCode = 0x14, Width = 124, IsCenterAligned = true });
            r3.Add(new KeyModel { English = "a", EnglishUpper = "A", Zhuyin = "ㄇ", VkCode = 0x41 });
            r3.Add(new KeyModel { English = "s", EnglishUpper = "S", Zhuyin = "ㄋ", VkCode = 0x53 });
            r3.Add(new KeyModel { English = "d", EnglishUpper = "D", Zhuyin = "ㄎ", VkCode = 0x44 });
            r3.Add(new KeyModel { English = "f", EnglishUpper = "F", Zhuyin = "ㄑ", VkCode = 0x46 });
            r3.Add(new KeyModel { English = "g", EnglishUpper = "G", Zhuyin = "ㄕ", VkCode = 0x47 });
            r3.Add(new KeyModel { English = "h", EnglishUpper = "H", Zhuyin = "ㄘ", VkCode = 0x48 });
            r3.Add(new KeyModel { English = "j", EnglishUpper = "J", Zhuyin = "ㄨ", VkCode = 0x4A });
            r3.Add(new KeyModel { English = "k", EnglishUpper = "K", Zhuyin = "ㄜ", VkCode = 0x4B });
            r3.Add(new KeyModel { English = "l", EnglishUpper = "L", Zhuyin = "ㄠ", VkCode = 0x4C });
            r3.Add(new KeyModel { English = ";", EnglishUpper = ":", Zhuyin = "ㄤ", VkCode = 0xBA });
            r3.Add(new KeyModel { English = "'", EnglishUpper = "\"", Zhuyin = "‘", VkCode = 0xDE });
            r3.Add(new KeyModel { English = "⏎", EnglishUpper = "Enter", Zhuyin = "", VkCode = 0x0D, Width = 105, IsCenterAligned = true });
            KeyRows.Add(r3);

            var r4 = new ObservableCollection<KeyModel>();
            r4.Add(new KeyModel { English = "⇧", EnglishUpper = "Shift", Zhuyin = "", VkCode = 0x10, Width = 164, IsCenterAligned = true });
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
            r4.Add(new KeyModel { English = "↑", EnglishUpper = "↑", Zhuyin = "", VkCode = 0x26, IsCenterAligned = true });
            r4.Add(new KeyModel { English = "⌨", EnglishUpper = "Fn", Zhuyin = "", VkCode = FN_KEY_CODE, IsCenterAligned = true });
            KeyRows.Add(r4);

            var r5 = new ObservableCollection<KeyModel>();
            r5.Add(new KeyModel { English = "⌃", EnglishUpper = "Ctrl", Zhuyin = "", VkCode = 0x11, Width = 85, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⊞", EnglishUpper = "Win", Zhuyin = "", VkCode = 0x5B, Width = 85, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⌥", EnglishUpper = "Alt", Zhuyin = "", VkCode = 0x12, Width = 85, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⎵", EnglishUpper = "Space", Zhuyin = "", VkCode = 0x20, Width = 450, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "Mode", EnglishUpper = "Mode", Zhuyin = "", VkCode = MODE_KEY_CODE, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "←", EnglishUpper = "←", Zhuyin = "", VkCode = 0x25, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "↓", EnglishUpper = "↓", Zhuyin = "", VkCode = 0x28, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "→", EnglishUpper = "→", Zhuyin = "", VkCode = 0x27, IsCenterAligned = true });
            KeyRows.Add(r5);

            foreach (var row in KeyRows)
            {
                foreach (var k in row)
                {
                    if (FnDisplayMap.TryGetValue(k.VkCode, out string? fnLabel) && !string.IsNullOrEmpty(fnLabel))
                    {
                        k.FnText = fnLabel;
                    }
                }
            }

            UpdateDisplay();
        }

        private void SetupFullKeyboard()
        {
            // F 鍵列 (F1-F12 + PrtSc, Scroll Lock, Pause)
            var r0 = new ObservableCollection<KeyModel>();
            r0.Add(new KeyModel { English = "Esc", EnglishUpper = "Esc", Zhuyin = "", VkCode = 0x1B, IsCenterAligned = true });

            r0.Add(new KeyModel { English = "F1", EnglishUpper = "F1", Zhuyin = "", VkCode = 0x70, LeftMargin = 66, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F2", EnglishUpper = "F2", Zhuyin = "", VkCode = 0x71, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F3", EnglishUpper = "F3", Zhuyin = "", VkCode = 0x72, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F4", EnglishUpper = "F4", Zhuyin = "", VkCode = 0x73, IsCenterAligned = true });

            r0.Add(new KeyModel { English = "F5", EnglishUpper = "F5", Zhuyin = "", VkCode = 0x74, LeftMargin = 32, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F6", EnglishUpper = "F6", Zhuyin = "", VkCode = 0x75, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F7", EnglishUpper = "F7", Zhuyin = "", VkCode = 0x76, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F8", EnglishUpper = "F8", Zhuyin = "", VkCode = 0x77, IsCenterAligned = true });

            r0.Add(new KeyModel { English = "F9", EnglishUpper = "F9", Zhuyin = "", VkCode = 0x78, LeftMargin = 32, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F10", EnglishUpper = "F10", Zhuyin = "", VkCode = 0x79, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F11", EnglishUpper = "F11", Zhuyin = "", VkCode = 0x7A, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "F12", EnglishUpper = "F12", Zhuyin = "", VkCode = 0x7B, IsCenterAligned = true });

            r0.Add(new KeyModel { English = "PrtSc", EnglishUpper = "PrtSc", Zhuyin = "", VkCode = 0x2C, LeftMargin = 32, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "ScrLk", EnglishUpper = "ScrLk", Zhuyin = "", VkCode = 0x91, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "Pause", EnglishUpper = "Pause", Zhuyin = "", VkCode = 0x13, IsCenterAligned = true });

            // 音量控制 (位於 Numpad 上方)
            r0.Add(new KeyModel { English = "🔇", EnglishUpper = "🔇", Zhuyin = "", VkCode = 0xAD, LeftMargin = 32, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "🔉", EnglishUpper = "🔉", Zhuyin = "", VkCode = 0xAE, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "🔊", EnglishUpper = "🔊", Zhuyin = "", VkCode = 0xAF, IsCenterAligned = true });
            r0.Add(new KeyModel { English = "⌃⌥⌦", EnglishUpper = "⌃⌥⌦", Zhuyin = "", VkCode = CAD_KEY_CODE, IsCenterAligned = true });

            KeyRows.Add(r0);

            // 數字列
            var r1 = new ObservableCollection<KeyModel>();
            r1.Add(new KeyModel { English = "`", EnglishUpper = "~", Zhuyin = "", VkCode = 0xC0 });
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
            r1.Add(new KeyModel { English = "⌫", EnglishUpper = "⌫", Zhuyin = "", VkCode = 0x08, Width = 120, IsCenterAligned = true });

            // 插入、Home、PageUp
            r1.Add(new KeyModel { English = "Ins", EnglishUpper = "Ins", Zhuyin = "", VkCode = 0x2D, LeftMargin = 32, IsCenterAligned = true });
            r1.Add(new KeyModel { English = "Home", EnglishUpper = "Home", Zhuyin = "", VkCode = 0x24, IsCenterAligned = true });
            r1.Add(new KeyModel { English = "PgUp", EnglishUpper = "PgUp", Zhuyin = "", VkCode = 0x21, IsCenterAligned = true });

            // Numpad 首列
            r1.Add(new KeyModel { English = "Num", EnglishUpper = "Num", Zhuyin = "", VkCode = 0x90, LeftMargin = 32, IsCenterAligned = true });
            r1.Add(new KeyModel { English = "/", EnglishUpper = "/", Zhuyin = "", VkCode = 0x6F, IsCenterAligned = true });
            r1.Add(new KeyModel { English = "*", EnglishUpper = "*", Zhuyin = "", VkCode = 0x6A, IsCenterAligned = true });
            r1.Add(new KeyModel { English = "-", EnglishUpper = "-", Zhuyin = "", VkCode = 0x6D, IsCenterAligned = true });
            KeyRows.Add(r1);

            var r2 = new ObservableCollection<KeyModel>();
            r2.Add(new KeyModel { English = "⇥", EnglishUpper = "Tab", Zhuyin = "", VkCode = 0x09, Width = 94, IsCenterAligned = true });
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
            r2.Add(new KeyModel { English = "[", EnglishUpper = "{", Zhuyin = "「", VkCode = 0xDB });
            r2.Add(new KeyModel { English = "]", EnglishUpper = "}", Zhuyin = "」", VkCode = 0xDD });
            r2.Add(new KeyModel { English = "\\", EnglishUpper = "|", Zhuyin = "、", VkCode = 0xDC, Width = 90 });

            // 刪除、End、PageDown
            r2.Add(new KeyModel { English = "Del", EnglishUpper = "Del", Zhuyin = "", VkCode = 0x2E, LeftMargin = 32, IsCenterAligned = true });
            r2.Add(new KeyModel { English = "End", EnglishUpper = "End", Zhuyin = "", VkCode = 0x23, IsCenterAligned = true });
            r2.Add(new KeyModel { English = "PgDn", EnglishUpper = "PgDn", Zhuyin = "", VkCode = 0x22, IsCenterAligned = true });

            // Numpad 7, 8, 9, Add
            r2.Add(new KeyModel { English = "7", EnglishUpper = "Home", Zhuyin = "", VkCode = 0x67, LeftMargin = 32 });
            r2.Add(new KeyModel { English = "8", EnglishUpper = "↑", Zhuyin = "", VkCode = 0x68 });
            r2.Add(new KeyModel { English = "9", EnglishUpper = "PgUp", Zhuyin = "", VkCode = 0x69 });
            r2.Add(new KeyModel { English = "+", EnglishUpper = "+", Zhuyin = "", VkCode = 0x6B, Height = 114, IsCenterAligned = true });
            KeyRows.Add(r2);

            var r3 = new ObservableCollection<KeyModel>();
            r3.Add(new KeyModel { English = "⇪", EnglishUpper = "Caps", Zhuyin = "", VkCode = 0x14, Width = 112, IsCenterAligned = true });
            r3.Add(new KeyModel { English = "a", EnglishUpper = "A", Zhuyin = "ㄇ", VkCode = 0x41 });
            r3.Add(new KeyModel { English = "s", EnglishUpper = "S", Zhuyin = "ㄋ", VkCode = 0x53 });
            r3.Add(new KeyModel { English = "d", EnglishUpper = "D", Zhuyin = "ㄎ", VkCode = 0x44 });
            r3.Add(new KeyModel { English = "f", EnglishUpper = "F", Zhuyin = "ㄑ", VkCode = 0x46 });
            r3.Add(new KeyModel { English = "g", EnglishUpper = "G", Zhuyin = "ㄕ", VkCode = 0x47 });
            r3.Add(new KeyModel { English = "h", EnglishUpper = "H", Zhuyin = "ㄘ", VkCode = 0x48 });
            r3.Add(new KeyModel { English = "j", EnglishUpper = "J", Zhuyin = "ㄨ", VkCode = 0x4A });
            r3.Add(new KeyModel { English = "k", EnglishUpper = "K", Zhuyin = "ㄜ", VkCode = 0x4B });
            r3.Add(new KeyModel { English = "l", EnglishUpper = "L", Zhuyin = "ㄠ", VkCode = 0x4C });
            r3.Add(new KeyModel { English = ";", EnglishUpper = ":", Zhuyin = "ㄤ", VkCode = 0xBA });
            r3.Add(new KeyModel { English = "'", EnglishUpper = "\"", Zhuyin = "‘", VkCode = 0xDE });
            r3.Add(new KeyModel { English = "⏎", EnglishUpper = "Enter", Zhuyin = "", VkCode = 0x0D, Width = 140, IsCenterAligned = true });

            // Numpad 4, 5, 6
            r3.Add(new KeyModel { English = "4", EnglishUpper = "←", Zhuyin = "", VkCode = 0x64, LeftMargin = 270 });
            r3.Add(new KeyModel { English = "5", EnglishUpper = "Clr", Zhuyin = "", VkCode = 0x65 });
            r3.Add(new KeyModel { English = "6", EnglishUpper = "→", Zhuyin = "", VkCode = 0x66 });
            KeyRows.Add(r3);

            var r4 = new ObservableCollection<KeyModel>();
            r4.Add(new KeyModel { English = "⇧", EnglishUpper = "Shift", Zhuyin = "", VkCode = 0x10, Width = 138, IsCenterAligned = true });
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
            r4.Add(new KeyModel { English = "⇧", EnglishUpper = "Shift", Zhuyin = "", VkCode = 0xA1, Width = 182, IsCenterAligned = true }); // 右Shift

            // ↑ 方向鍵
            r4.Add(new KeyModel { English = "↑", EnglishUpper = "↑", Zhuyin = "", VkCode = 0x26, LeftMargin = 102, IsCenterAligned = true });

            // Numpad 1, 2, 3, Enter
            r4.Add(new KeyModel { English = "1", EnglishUpper = "End", Zhuyin = "", VkCode = 0x61, LeftMargin = 102 });
            r4.Add(new KeyModel { English = "2", EnglishUpper = "↓", Zhuyin = "", VkCode = 0x62 });
            r4.Add(new KeyModel { English = "3", EnglishUpper = "PgDn", Zhuyin = "", VkCode = 0x63 });
            r4.Add(new KeyModel { English = "⏎", EnglishUpper = "Enter", Zhuyin = "", VkCode = 0x0D, Height = 114, IsCenterAligned = true });
            KeyRows.Add(r4);

            var r5 = new ObservableCollection<KeyModel>();
            r5.Add(new KeyModel { English = "⌃", EnglishUpper = "Ctrl", Zhuyin = "", VkCode = 0x11, Width = 90, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⊞", EnglishUpper = "Win", Zhuyin = "", VkCode = 0x5B, Width = 75, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⌥", EnglishUpper = "Alt", Zhuyin = "", VkCode = 0x12, Width = 75, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⎵", EnglishUpper = "Space", Zhuyin = "", VkCode = 0x20, Width = 446, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⌥", EnglishUpper = "Alt", Zhuyin = "", VkCode = 0xA5, Width = 75, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⊞", EnglishUpper = "Win", Zhuyin = "", VkCode = 0x5B, Width = 75, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "Mode", EnglishUpper = "Mode", Zhuyin = "", VkCode = MODE_KEY_CODE, Width = 75, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "⌃", EnglishUpper = "Ctrl", Zhuyin = "", VkCode = 0xA3, Width = 75, IsCenterAligned = true });

            // ←, ↓, →
            r5.Add(new KeyModel { English = "←", EnglishUpper = "←", Zhuyin = "", VkCode = 0x25, LeftMargin = 32, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "↓", EnglishUpper = "↓", Zhuyin = "", VkCode = 0x28, IsCenterAligned = true });
            r5.Add(new KeyModel { English = "→", EnglishUpper = "→", Zhuyin = "", VkCode = 0x27, IsCenterAligned = true });

            // Numpad 0, .
            r5.Add(new KeyModel { English = "0", EnglishUpper = "Ins", Zhuyin = "", VkCode = 0x60, Width = 136, LeftMargin = 32 });
            r5.Add(new KeyModel { English = ".", EnglishUpper = "Del", Zhuyin = "", VkCode = 0x6E });
            KeyRows.Add(r5);

            UpdateDisplay();
        }

        #region INI 設定檔讀寫 (透明度, 視窗位置, 自訂按鍵版面)
        private void LoadSettings()
        {
            bool hasSavedTheme = false;
            // [修正] 即使文件不存在，也應執行後續的主題偵測等邏輯
            if (File.Exists(_iniFilePath))
            {
                try
                {
                var lines = File.ReadAllLines(_iniFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('=', 2);
                    if (parts.Length != 2) continue;

                    string key = parts[0].Trim();
                    string val = parts[1].Trim();

                    if (key == "Opacity" && double.TryParse(val, out double op))
                    {
                        double finalOp = Math.Max(0.2, Math.Min(1.0, op));
                        this.Opacity = finalOp;
                        if (OpacitySlider != null) OpacitySlider.Value = finalOp;
                    }
                    else if (key == "Left" && double.TryParse(val, out double left))
                    {
                        TargetLeft = left;
                    }
                    else if (key == "Top" && double.TryParse(val, out double top))
                    {
                        TargetTop = top;
                    }
                    else if (key == "IsFullLayout" && bool.TryParse(val, out bool isFull))
                    {
                        if (this.IsFullLayout != isFull) ToggleFullLayout(null);
                    }
                    else if (key == "IsDynamicLayout" && bool.TryParse(val, out bool isDyn))
                    {
                        this.IsDynamicLayout = isDyn;
                    }
                    else if (key == "IsPinned" && bool.TryParse(val, out bool isPin))
                    {
                        _isPinned = isPin;
                        this.Topmost = isPin;
                        OnPropertyChanged("PinIcon");
                    }
                    else if (key == "Width" && double.TryParse(val, out double w))
                    {
                        TargetWidth = w;
                    }

                    else if (key == "Height" && double.TryParse(val, out double h))
                    {
                        TargetHeight = h;
                    }
                    else if (key == "IsDarkMode" && bool.TryParse(val, out bool isDark))
                    {
                        hasSavedTheme = true;
                        ApplyTheme(isDark);
                    }
                    else if (key == "BgColor" && !string.IsNullOrEmpty(val))
                    {
                        // 舊版相容性
                    }
                    else if (key == "TextColor" && !string.IsNullOrEmpty(val))
                    {
                        // 舊版相容性
                    }
                }

                // 嘗試套用自訂排列
                ApplyCustomLayout();
            }
            catch { }
            }

            if (!hasSavedTheme)
            {
                DetectSystemTheme();
            }
        }

        private void ApplyCustomLayout()
        {
            if (!File.Exists(_iniFilePath)) return;
            string prefix = IsFullLayout ? "FullRow" : "CompactRow";

            try
            {
                var lines = File.ReadAllLines(_iniFilePath);
                var dict = new Dictionary<string, string>();
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var parts = line.Split('=', 2);
                    if (parts.Length == 2) dict[parts[0].Trim()] = parts[1].Trim();
                }

                // 若該版面設定不存在，直接返回
                bool hasLayout = false;
                for (int i = 0; i < KeyRows.Count; i++)
                {
                    if (dict.ContainsKey($"{prefix}{i}")) hasLayout = true;
                }
                if (!hasLayout) return;

                // 收集所有「原始的 KeyModel」供跨列全域搜尋
                var allOriginalKeys = KeyRows.SelectMany(r => r).ToList();
                var keyOriginalRowMap = new Dictionary<KeyModel, int>();
                var originalLayoutsPerRow = new List<List<(double Width, double Height, double LeftMargin, double TopMargin, bool IsCenterAligned)>>();

                for (int i = 0; i < KeyRows.Count; i++)
                {
                    var layouts = new List<(double Width, double Height, double LeftMargin, double TopMargin, bool IsCenterAligned)>();
                    foreach (var k in KeyRows[i])
                    {
                        keyOriginalRowMap[k] = i;
                        layouts.Add((k.Width, k.Height, k.LeftMargin, k.TopMargin, k.IsCenterAligned));
                    }
                    originalLayoutsPerRow.Add(layouts);
                }

                var newRows = new List<ObservableCollection<KeyModel>>();
                for (int i = 0; i < KeyRows.Count; i++)
                {
                    newRows.Add(new ObservableCollection<KeyModel>());
                }

                // 根據 ini 把各按鈕放進目標的列與位置
                for (int i = 0; i < KeyRows.Count; i++)
                {
                    string layoutKey = $"{prefix}{i}";
                    if (dict.TryGetValue(layoutKey, out string? orderStr) && !string.IsNullOrEmpty(orderStr))
                    {
                        var ids = orderStr.Split(',').ToList();
                        foreach (var id in ids)
                        {
                            var match = allOriginalKeys.FirstOrDefault(k => k.Id == id);
                            if (match != null)
                            {
                                newRows[i].Add(match);
                                allOriginalKeys.Remove(match); // 移出總清單，表示已被用掉
                            }
                        }
                    }
                }

                // 把設定檔沒記載的剩餘按鍵塞回預設預期會出現的列，補齊版面
                foreach (var remain in allOriginalKeys)
                {
                    newRows[keyOriginalRowMap[remain]].Add(remain);
                }

                // 將實體排版重新套用到底層的新順序上，讓按鈕的尺寸與間距保留在原位不崩壞
                for (int i = 0; i < KeyRows.Count; i++)
                {
                    var newRow = newRows[i];
                    var layouts = originalLayoutsPerRow[i];

                    for (int j = 0; j < newRow.Count && j < layouts.Count; j++)
                    {
                        newRow[j].Width = layouts[j].Width;
                        newRow[j].Height = layouts[j].Height;
                        newRow[j].LeftMargin = layouts[j].LeftMargin;
                        newRow[j].TopMargin = layouts[j].TopMargin;
                        newRow[j].IsCenterAligned = layouts[j].IsCenterAligned;
                    }

                    KeyRows[i] = newRow;
                }
            }
            catch { }
        }

        private void SaveSettings()
        {
            try
            {
                double op = this.Opacity;
                double w = TargetWidth;
                double h = TargetHeight;
                double left = TargetLeft;
                double top = TargetTop;

                // [修正] 寫入前檢查數值有效性，防止 NaN 或 Infinity 損壞設定檔
                if (!double.IsFinite(op)) op = 1.0;
                if (!double.IsFinite(w) || w < 100) w = 1000;
                if (!double.IsFinite(h) || h < 50) h = 300;
                if (!double.IsFinite(left)) left = 0;
                if (!double.IsFinite(top)) top = 0;

                var lines = new List<string>
                {
                    $"Opacity={op}",
                    $"Width={w}",
                    $"Height={h}",
                    $"Left={left}",
                    $"Top={top}",
                    $"IsFullLayout={this.IsFullLayout}",
                    $"IsDynamicLayout={this.IsDynamicLayout}",
                    $"IsPinned={_isPinned}",
                    $"IsDarkMode={_isDarkMode}"
                };

                // 儲存目前的版面按鍵順序
                string prefix = IsFullLayout ? "FullRow" : "CompactRow";
                for (int i = 0; i < KeyRows.Count; i++)
                {
                    var ids = KeyRows[i].Select(k => k.Id);
                    lines.Add($"{prefix}{i}={string.Join(",", ids)}");
                }

                // 也讀取未選上版面的舊紀錄，避免被洗掉
                if (File.Exists(_iniFilePath))
                {
                    var oldLines = File.ReadAllLines(_iniFilePath);
                    string otherPrefix = IsFullLayout ? "CompactRow" : "FullRow";
                    foreach (var oldLine in oldLines)
                    {
                        if (oldLine.StartsWith(otherPrefix))
                        {
                            lines.Add(oldLine);
                        }
                    }
                }

                File.WriteAllLines(_iniFilePath, lines);
            }
            catch { }
        }
        #endregion

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            Task.Delay(500).ContinueWith(_ => this.Dispatcher.Invoke(() => ReduceMemoryUsage()));
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "確定要關閉虛擬鍵盤嗎？",
                "OSK 關閉確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);

            if (result == MessageBoxResult.Yes)
            {
                System.Windows.Application.Current.Shutdown();
            }
        }

        private DispatcherTimer? _oskMonitorTimer = null;
        private bool _savedPinnedStateBeforeSystemOsk = true;

        private void CallSystemOsk_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // [需求] 備份目前的釘選狀態，並強制設為固定 (Pinned)
                _savedPinnedStateBeforeSystemOsk = _isPinned;
                if (!_isPinned)
                {
                    _isPinned = true;
                    OnPropertyChanged("PinIcon");
                }

                // 為了防止我們的 OSK 以 32 位元執行時，呼叫 OSK 遭到 WOW64 轉向導致失敗
                // 我們強制指定使用 64 位元系統的原生資料夾 (sysnative)
                string sysDir = Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess
                    ? Path.Combine(Environment.GetEnvironmentVariable("windir") ?? @"C:\Windows", "sysnative")
                    : Environment.GetFolderPath(Environment.SpecialFolder.System);

                string oskPath = Path.Combine(sysDir, "osk.exe");

                // [重要] 在啟動系統小鍵盤前，必須先停止本體的 UIA 偵測器
                // 防止兩套 UIA 掛鉤同時運作造成資源衝突、卡死或效能嚴重下降
                _inputDetector?.Stop();

                // [優化] 改用 cmd /c start 方式啟動，這在某些系統權限下比直接 Process.Start 更穩定
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c start \"\" \"{oskPath}\"",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true,
                    UseShellExecute = false
                });

                // 設定系統 OSK 啟用的全域旗標，這將會使得 IsBusySuppressing 成立，從而阻擋常駐偵測器干涉
                _isSystemOskActive = true;

                // 隱藏本體，避免畫面重疊
                this.Hide();

                int currentProcessId = Process.GetCurrentProcess().Id;
                DateTime startTime = DateTime.Now;
                bool didSystemOskStart = false; // 紀錄是否曾經偵測到系統鍵盤啟動

                // 啟動監控，等待系統 osk.exe 關閉後再重新顯示本體
                if (_oskMonitorTimer == null)
                {
                    _oskMonitorTimer = new DispatcherTimer();
                    _oskMonitorTimer.Interval = TimeSpan.FromSeconds(1);
                    _oskMonitorTimer.Tick += (s, args) =>
                    {
                        var procs = Process.GetProcessesByName("osk");
                        
                        // [雙重辨認] 1. 進程路徑比對 2. FindWindow 視窗類別比對 (解決權限導致無法讀取進程資訊的邊界案例)
                        bool hasOskProcess = procs.Any(p => {
                            try {
                                string? path = p.MainModule?.FileName?.ToLower();
                                if (string.IsNullOrEmpty(path)) return false;
                                return (path.Contains(@"\windows\system32\") || path.Contains(@"\windows\sysnative\")) && p.Id != currentProcessId;
                            } catch { 
                                return p.Id != currentProcessId && p.ProcessName.Equals("osk", StringComparison.OrdinalIgnoreCase); 
                            } 
                        });

                        // 系統 OSK 的視窗類別名稱固定為 OSKMainClass
                        bool hasOskWindow = FindWindow("OSKMainClass", null) != IntPtr.Zero;

                        bool isSystemOskRunning = hasOskProcess || hasOskWindow;

                        if (isSystemOskRunning) didSystemOskStart = true;

                        double elapsed = (DateTime.Now - startTime).TotalSeconds;
                        bool initialBuffer = elapsed < 3;
                        
                        // [邏輯優化] 只有在「從未偵測到啟動」且「超過 15 秒」時才判定為真的啟動失敗
                        // 一旦偵測到啟動 (didSystemOskStart = true)，則永遠等待直到它消失為止，不再受超時限制
                        bool launchTimeout = !didSystemOskStart && elapsed > 15;

                        if ((!isSystemOskRunning && !initialBuffer && didSystemOskStart) || launchTimeout)
                        {
                            _isSystemOskActive = false; // 取消抑制標記
                            
                            // [關鍵修復] 使用統一方法恢復原本的釘選狀態
                            ApplyPinState(_savedPinnedStateBeforeSystemOsk);
                            OnPropertyChanged("PinIcon");

                            // [恢復] 重新啟動偵測器 (僅在非釘選狀態下啟動)
                            if (!_isPinned) _inputDetector?.Start();
                            
                            this.Show();
                            _oskMonitorTimer?.Stop();
                            _oskMonitorTimer = null;
                            
                            if (launchTimeout)
                            {
                                System.Windows.MessageBox.Show("內建鍵盤未能在預期時間內啟動，已恢復本體。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                    };
                }
                _oskMonitorTimer.Start();
            }
            catch (Exception ex)
            {
                _isSystemOskActive = false;
                ApplyPinState(_savedPinnedStateBeforeSystemOsk);
                this.Show();
                System.Windows.MessageBox.Show($"無法啟動內建虛擬鍵盤 [{ex.Message}]", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RestoreSize_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "確定要恢復預設配置嗎？\n您將會遺失所有自定義的按鍵位置與透明度設定。",
                "重置警告",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (result != MessageBoxResult.Yes) return;

            // [效能優化] 移除冗餘的 ToggleFullLayout 呼叫，直接在下方統一處理初始化
            _isFullLayout = false; // 強制設為精簡版面
            FullLayoutIcon = "⌨";

            TargetWidth = 1000;
            double ratio = 0.31;
            double minH = Math.Max(TargetWidth * ratio + 45, 120);
            TargetHeight = minH;

            this.Opacity = 1.0;
            this.IsEditMode = false;

            // [關鍵修復] 使用統一方法重置圖釘與佈局狀態
            ApplyPinState(true);
            ApplyLayoutType(true);

            // [新增] 重置物理特效
            _isDragFxEnabled = true;
            OnPropertyChanged("DragFxIcon");

            // 恢復顏色為系統預設
            DetectSystemTheme();

            if (OpacitySlider != null) OpacitySlider.Value = 1.0;
            MainRootGrid.Opacity = 1.0;

            if (File.Exists(_iniFilePath))
            {
                try { File.Delete(_iniFilePath); } catch { }
            }

            // [效能優化] 統一執行一次清空與建立
            KeyRows.Clear();
            SetupKeyboard();

            KeyBoardItemsControl.ItemsSource = null;
            KeyBoardItemsControl.ItemsSource = KeyRows;

            // 讓視窗回到中央位置
            CenterWindow();
        }

        private bool _isMouseDraggingWindow = false;
        private System.Windows.Point _mouseDragStartPoint;

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null || _isTouchDraggingWindow) return;
            if (_isResizing) return;

            // Manual dragging logic for VirtualCanvas
            _isMouseDraggingWindow = true;
            _mouseDragStartPoint = e.GetPosition(this);
            Mouse.Capture((UIElement)sender);
            e.Handled = true;
        }


        private void TopTitleBar_TouchDown(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (IsEditMode) return;
            if (sender is UIElement el)
            {
                _isTouchDraggingWindow = true;
                // 記錄在視窗中的絕對位置，避免隨視窗移動而產生無限遞迴的偏移
                _touchDragStartPoint = e.GetTouchPoint(this).Position;
                el.CaptureTouch(e.TouchDevice);
                e.Handled = true; 
            }
        }

        private void TopTitleBar_TouchMove(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (_isTouchDraggingWindow)
            {
                var pos = e.GetTouchPoint(this).Position;
                double dx = pos.X - _touchDragStartPoint.X;
                double dy = pos.Y - _touchDragStartPoint.Y;

                TargetLeft += dx;
                TargetTop += dy;
                
                // 由於 TargetLeft/Top 已經改變，視窗內容也移動了，下次取得的相對 pos 其實會跟著跑。
                // 因此我們需要不斷更新 _touchDragStartPoint
                _touchDragStartPoint = pos;
            }
        }

        private void Window_ManipulationBoundaryFeedback(object sender, ManipulationBoundaryFeedbackEventArgs e)
        {
            e.Handled = true;
        }

        private void TopTitleBar_TouchUp(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (_isTouchDraggingWindow)
            {
                _isTouchDraggingWindow = false;
                if (sender is UIElement el) el.ReleaseTouchCapture(e.TouchDevice);
                ClampToScreen(); // 確保不會被拖出畫面外
                SaveSettings();  // 拖曳結束後存檔
            }
        }
        private void Resize_Init(object sender, MouseButtonEventArgs e) { _isResizing = true; Mouse.Capture((UIElement)sender); }
        private void Resize_End(object sender, MouseButtonEventArgs e) { _isResizing = false; Mouse.Capture(null); }
        private void Resize_Move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isResizing)
            {
                // 必須相對於 KeyboardContainer 取得滑鼠位置，而非相對於全螢幕的透明 Window
                // 否則 p.X 可能高達數千像素，使鍵盤瞬間放大到全螢幕
                System.Windows.Point p = e.GetPosition(KeyboardContainer);
                if (p.X > 200) TargetWidth = p.X;

                double ratio = _isFullLayout ? 0.23 : 0.31;
                double minH = Math.Max(TargetWidth * ratio + 5, 80);
                TargetHeight = minH;
            }
        }


        protected override void OnClosed(EventArgs e)
        {
            System.Windows.Media.CompositionTarget.Rendering -= VisualSyncTimer_Tick;
            _notifyIcon?.Dispose();
            
            // [快速關閉] 傳入 true 以跳過容易掛死的 UIA 事件註銷，由 OS 統一回收
            _inputDetector?.Stop(true); 
            _memoryTrimTimer?.Stop();
            _imeTimer?.Stop();
            _oskMonitorTimer?.Stop();
            
            if (!IsEditMode) SaveSettings();
            base.OnClosed(e);

            // [最終補強] 徹底強制終止進程，不留任何背景殭屍執行緒，確保可立即重複啟動測試
            Environment.Exit(0);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            if (_virtualAltToggle)
            {
                SendSimulatedKey(0x12, true);
                _virtualAltToggle = false;
            }
            base.OnClosing(e);
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        protected void OnPropertyChanged(PropertyChangedEventArgs args) => PropertyChanged?.Invoke(this, args);

        private void ShowSecurityMenu()
        {
            this.Dispatcher.Invoke(() =>
            {
                var menu = new System.Windows.Controls.ContextMenu();

                var miLock = new System.Windows.Controls.MenuItem { Header = "鎖定" };
                miLock.Click += (s, e) => { LockWorkStation(); };

                var miSwitch = new System.Windows.Controls.MenuItem { Header = "切換使用者" };
                miSwitch.Click += (s, e) => { StartSwitchUser(); };

                var miSignOut = new System.Windows.Controls.MenuItem { Header = "登出" };
                miSignOut.Click += (s, e) => { ExitWindowsEx(0x00000000u, 0); };

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
                SendSimulatedKey(0x11, false);
                SendSimulatedKey(0x10, false);
                SendSimulatedKey(0x1B, false);
                SendSimulatedKey(0x1B, true);
                SendSimulatedKey(0x10, true);
                SendSimulatedKey(0x11, true);
            }
        }

        private void StartSwitchUser()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("tsdiscon") { UseShellExecute = true });
            }
            catch
            {
                LockWorkStation();
            }
        }

        private void ShowChangePassword()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("control", "/name Microsoft.UserAccounts") { UseShellExecute = true });
            }
            catch { }
        }
    }

    public class KeyModel : INotifyPropertyChanged
    {
        private string _english = "";
        public string English { get { return _english; } set { if (_english != value) { _english = value; OnPropertyChanged("English"); OnPropertyChanged("EnglishDisp"); } } }

        private string _englishUpper = "";
        public string EnglishUpper { get { return _englishUpper; } set { if (_englishUpper != value) { _englishUpper = value; OnPropertyChanged("EnglishUpper"); OnPropertyChanged("ShiftDisp"); } } }

        private string _zhuyin = "";
        public string Zhuyin { get { return _zhuyin; } set { if (_zhuyin != value) { _zhuyin = value; OnPropertyChanged("Zhuyin"); OnPropertyChanged("ZhuyinDisp"); } } }

        private string _fnText = "";
        public string FnText { get { return _fnText; } set { if (_fnText != value) { _fnText = value; OnPropertyChanged("FnText"); OnPropertyChanged("FnDisp"); } } }

        private string _displayName = "";
        public string DisplayName { get { return _displayName; } set { if (_displayName != value) { _displayName = value; OnPropertyChanged(_argsDisplayName); } } }

        private string _dynamicTextColor = "White";
        public string DynamicTextColor { get { return _dynamicTextColor; } set { if (_dynamicTextColor != value) { _dynamicTextColor = value; OnPropertyChanged(_argsDynamicTextColor); } } }

        public byte VkCode { get; set; }
        private static readonly Random _rnd = new();

        // 動態隨動屬性
        private double _offsetX = 0;
        public double OffsetX { get => _offsetX; set { if (Math.Abs(_offsetX - value) > 0.0001) { _offsetX = value; OnPropertyChanged(_argsOffsetX); } } }

        private double _offsetY = 0;
        public double OffsetY { get => _offsetY; set { if (Math.Abs(_offsetY - value) > 0.0001) { _offsetY = value; OnPropertyChanged(_argsOffsetY); } } }

        public double VelocityX { get; set; } = 0;
        public double VelocityY { get; set; } = 0;
        public double FollowSpeed { get; set; } = 0.2 + _rnd.NextDouble() * 0.3; // 0.2 ~ 0.5 (穩定追隨)
        public double Friction { get; set; } = 0.5 + _rnd.NextDouble() * 0.15;   // 強阻尼 (0.5 ~ 0.65) 消除擺動
        public double NoisePhase { get; set; } = _rnd.NextDouble() * Math.PI * 2;
        public double Mass { get; set; } = 1.0 + _rnd.NextDouble() * 0.5;         // 增加質量 (1.0 ~ 1.5) 提高穩定性

        public void RandomizePhysics()
        {
            FollowSpeed = 0.2 + _rnd.NextDouble() * 0.3;
            Friction = 0.5 + _rnd.NextDouble() * 0.15;
            Mass = 1.0 + _rnd.NextDouble() * 0.5;
            NoisePhase = _rnd.NextDouble() * Math.PI * 2;
        }

        private double _width = 65;
        public double Width
        {
            get => _width;
            set
            {
                if (_width != value)
                {
                    _width = value;
                    OnPropertyChanged(nameof(Width));
                }
            }
        }

        public string Id => $"{VkCode}_{System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(English ?? ""))}";

        private double _height = 55;
        public double Height
        {
            get => _height;
            set
            {
                if (_height != value)
                {
                    _height = value;
                    OnPropertyChanged(nameof(Height));
                }
            }
        }

        private double _leftMargin = 2;
        public double LeftMargin
        {
            get => _leftMargin;
            set
            {
                if (_leftMargin != value)
                {
                    _leftMargin = value;
                    OnPropertyChanged(nameof(LeftMargin));
                    OnPropertyChanged(nameof(Margin));
                }
            }
        }

        private double _topMargin = 2;
        public double TopMargin
        {
            get => _topMargin;
            set
            {
                if (_topMargin != value)
                {
                    _topMargin = value;
                    OnPropertyChanged(nameof(TopMargin));
                    OnPropertyChanged(nameof(Margin));
                }
            }
        }

        public System.Windows.Thickness Margin => new System.Windows.Thickness(_leftMargin, _topMargin, 2, _height > 55 ? 55 - _height + 2 : 2);

        private bool _isCenterAligned = false;
        public bool IsCenterAligned
        {
            get => _isCenterAligned;
            set
            {
                if (_isCenterAligned != value)
                {
                    _isCenterAligned = value;
                    OnPropertyChanged(nameof(IsCenterAligned));
                    OnPropertyChanged(nameof(MainColSpan));
                    OnPropertyChanged(nameof(MainRowSpan));
                    OnPropertyChanged(nameof(MainHorizontalAlignment));
                    OnPropertyChanged(nameof(MainVerticalAlignment));
                }
            }
        }

        public int MainColSpan => IsCenterAligned ? 2 : 1;
        public int MainRowSpan => IsCenterAligned ? 2 : 1;
        public string MainHorizontalAlignment => IsCenterAligned ? "Center" : "Left";
        public string MainVerticalAlignment => IsCenterAligned ? "Center" : "Top";

        public string EnglishDisp => string.IsNullOrEmpty(English) ? " " : English;
        public string ShiftDisp
        {
            get
            {
                if (string.IsNullOrEmpty(EnglishUpper) || EnglishUpper == English) return " ";
                bool isNumpad = VkCode >= 0x60 && VkCode <= 0x69 || VkCode == 0x6E;
                if (!isNumpad && EnglishUpper.Length > 1) return " ";
                return EnglishUpper;
            }
        }
        public string ZhuyinDisp => string.IsNullOrEmpty(Zhuyin) ? " " : Zhuyin;
        public string FnDisp => string.IsNullOrEmpty(FnText) ? " " : FnText;

        private string _textColor = "White";
        public string TextColor { get { return _textColor; } set { _textColor = value; OnPropertyChanged("TextColor"); } }

        private string _shiftColor = "#1E90FF";
        public string ShiftColor
        {
            get
            {
                bool isNumpad = VkCode >= 0x60 && VkCode <= 0x69 || VkCode == 0x6E;
                if (isNumpad && !string.IsNullOrEmpty(EnglishUpper) && EnglishUpper != English)
                    return _numColor;
                return _shiftColor;
            }
            set { _shiftColor = value; OnPropertyChanged("ShiftColor"); }
        }

        private string _fnColor = "#32CD32";
        public string FnColor { get { return _fnColor; } set { _fnColor = value; OnPropertyChanged("FnColor"); } }

        private string _zhuyinColor = "Orange";
        public string ZhuyinColor { get { return _zhuyinColor; } set { _zhuyinColor = value; OnPropertyChanged("ZhuyinColor"); } }

        private string _numColor = "LightSkyBlue";
        public string NumColor { get { return _numColor; } set { _numColor = value; OnPropertyChanged("NumColor"); OnPropertyChanged("ShiftColor"); } }

        private bool _isPressed = false;
        public bool IsPressed { get { return _isPressed; } set { if (_isPressed != value) { _isPressed = value; OnPropertyChanged(_argsIsPressed); OnPropertyChanged(_argsBackground); } } }

        private bool _isActiveToggle = false;
        public bool IsActiveToggle { get { return _isActiveToggle; } set { if (_isActiveToggle != value) { _isActiveToggle = value; OnPropertyChanged(_argsIsActiveToggle); OnPropertyChanged(_argsBackground); } } }

        private string _normalBackground = "#333333";
        private string _pressedBackground = "#666666";

        public void SetThemeColors(string normal, string pressed, string text, string shift, string fn, string zy, string num)
        {
            _normalBackground = normal;
            _pressedBackground = pressed;
            TextColor = text;
            ShiftColor = shift;
            FnColor = fn;
            ZhuyinColor = zy;
            NumColor = num;
            OnPropertyChanged("Background");
        }

        private string? _customBackground = null;
        public string Background
        {
            get
            {
                if (_customBackground != null) return _customBackground;
                if (_isPressed) return _pressedBackground;
                if (_isActiveToggle) return _pressedBackground;
                return _normalBackground;
            }
            set
            {
                _customBackground = value;
                OnPropertyChanged("Background");
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        protected void OnPropertyChanged(PropertyChangedEventArgs args) => PropertyChanged?.Invoke(this, args);

        private static readonly PropertyChangedEventArgs _argsOffsetX = new PropertyChangedEventArgs(nameof(OffsetX));
        private static readonly PropertyChangedEventArgs _argsOffsetY = new PropertyChangedEventArgs(nameof(OffsetY));
        private static readonly PropertyChangedEventArgs _argsIsPressed = new PropertyChangedEventArgs(nameof(IsPressed));
        private static readonly PropertyChangedEventArgs _argsBackground = new PropertyChangedEventArgs(nameof(Background));
        private static readonly PropertyChangedEventArgs _argsDisplayName = new PropertyChangedEventArgs(nameof(DisplayName));
        private static readonly PropertyChangedEventArgs _argsDynamicTextColor = new PropertyChangedEventArgs(nameof(DynamicTextColor));
        private static readonly PropertyChangedEventArgs _argsIsActiveToggle = new PropertyChangedEventArgs(nameof(IsActiveToggle));
    }

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
