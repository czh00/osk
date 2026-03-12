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
        // ---------------------------------

        [DllImport("user32.dll")] private static extern uint RegisterWindowMessage(string lpString);
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

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

        public ObservableCollection<ObservableCollection<KeyModel>> KeyRows { get; set; } = new();
        public ICommand? KeyCommand { get; set; }
        public ICommand? ToggleThemeCommand { get; set; }
        public ICommand? TogglePinCommand { get; set; }
        public ICommand? ToggleLayoutCommand { get; set; }
        public ICommand? ToggleFullLayoutCommand { get; set; }
        public ICommand? ToggleEditModeCommand { get; set; }

        private bool _isPinned = true;
        public string PinIcon { get { return _isPinned ? "📍" : "📌"; } }

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

        private string _fullLayoutIcon = "📱";
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

        // 視覺同步
        private DispatcherTimer? _visualSyncTimer;
        private int _imeCheckCounter = 0;

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
        #endregion

        // 將特定按鍵加入到 User32 SendInput 結構清單中
        private void AddKeyInput(List<INPUT> inputs, byte vk, bool keyUp)
        {
            uint flags = keyUp ? KEYEVENTF_KEYUP : 0;
            // 針對擴展鍵 (Extended Key) 加入 0x0001 旗標，確保 Windows 能正確識別左右鍵或特殊鍵 (如方向鍵、NumLock)
            if (vk == 0x90 || vk == 0x21 || vk == 0x22 || vk == 0x23 || vk == 0x24 ||
                vk == 0x25 || vk == 0x26 || vk == 0x27 || vk == 0x28 || vk == 0x2D ||
                vk == 0x2E || vk == 0xA1 || vk == 0xA3 || vk == 0xA5 || vk == 0x5B || vk == 0x5C || vk == 0x6F ||
                vk == 0xAD || vk == 0xAE || vk == 0xAF)
            {
                flags |= 0x0001;
            }

            var input = new INPUT
            {
                type = INPUT_KEYBOARD,
                u = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = vk,
                        wScan = 0,
                        dwFlags = flags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
            inputs.Add(input);
        }

        // 傳送單一按鍵的按下或放開事件
        private void SendSimulatedKey(byte vk, bool isKeyUp)
        {
            var list = new List<INPUT>();
            AddKeyInput(list, vk, isKeyUp);
            SendInput((uint)list.Count, list.ToArray(), INPUT.Size);
        }

        // 清除所有虛擬修飾鍵的暫存狀態
        private void ResetVirtualModifiers()
        {
            _virtualShiftToggle = false;
            _virtualCtrlToggle = false;
            _virtualWinToggle = false;
            _virtualAltToggle = false;
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
            ToggleLayoutCommand = new RelayCommand<object>(ToggleLayout);
            ToggleFullLayoutCommand = new RelayCommand<object>(ToggleFullLayout);
            ToggleEditModeCommand = new RelayCommand<object>(ToggleEditMode);

            SetupKeyboard();

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

            _visualSyncTimer = new DispatcherTimer(DispatcherPriority.Render);
            _visualSyncTimer.Interval = TimeSpan.FromMilliseconds(30);
            _visualSyncTimer.Tick += VisualSyncTimer_Tick;
            _visualSyncTimer.Start();
            _inputDetector = new InputDetector(this);
            if (!_isPinned)
            {
                _inputDetector.Start();
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

        private void TogglePin(object? parameter)
        {
            _isPinned = !_isPinned;
            OnPropertyChanged("PinIcon");

            if (_inputDetector != null)
            {
                if (_isPinned)
                    _inputDetector.Stop();
                else
                    _inputDetector.Start();
            }
        }

        private void ToggleLayout(object? parameter)
        {
            IsDynamicLayout = !IsDynamicLayout;
            LayoutIcon = IsDynamicLayout ? "🗔" : "🗖";
            UpdateDisplay();
        }

        private void ToggleFullLayout(object? parameter)
        {
            IsFullLayout = !IsFullLayout;
            FullLayoutIcon = IsFullLayout ? "📱" : "⌨";

            KeyRows.Clear();
            if (IsFullLayout)
            {
                SetupFullKeyboard();
                double targetH = Math.Max(this.Width * 0.23 + 5, 80);
                this.MinHeight = targetH;
                this.MaxHeight = targetH;
                this.Height = targetH;
            }
            else
            {
                SetupKeyboard();
                double targetH = Math.Max(this.Width * 0.31 + 5, 80);
                this.MinHeight = targetH;
                this.MaxHeight = targetH;
                this.Height = targetH;
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

                var source = PresentationSource.FromVisual(this);
                Matrix transform = Matrix.Identity;
                if (source?.CompositionTarget != null) transform = source.CompositionTarget.TransformFromDevice;

                var topLeft = transform.Transform(new System.Windows.Point(wa.Left, wa.Top));
                var bottomRight = transform.Transform(new System.Windows.Point(wa.Right, wa.Bottom));
                double waLeft = topLeft.X;
                double waTop = topLeft.Y;
                double waWidth = Math.Max(0, bottomRight.X - topLeft.X);
                double waHeight = Math.Max(0, bottomRight.Y - topLeft.Y);

                double w = this.ActualWidth > 0 ? this.ActualWidth : (double.IsNaN(this.Width) ? 1000 : this.Width);
                double h = this.ActualHeight > 0 ? this.ActualHeight : (double.IsNaN(this.Height) ? 360 : this.Height);

                double newLeft = this.Left;
                double newTop = this.Top;

                if (newLeft < waLeft) newLeft = waLeft;
                if (newLeft + w > waLeft + waWidth) newLeft = waLeft + waWidth - w;
                if (newTop < waTop) newTop = waTop;
                if (newTop + h > waTop + waHeight) newTop = waTop + waHeight - h;

                if (this.Left != newLeft) this.Left = newLeft;
                if (this.Top != newTop) this.Top = newTop;
            }
            finally
            {
                _isClamping = false;
            }
        }

        private void CenterWindow()
        {
            var cursorPos = System.Windows.Forms.Cursor.Position;
            var screen = System.Windows.Forms.Screen.FromPoint(cursorPos);
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

            double w = this.ActualWidth > 0 ? this.ActualWidth : (double.IsNaN(this.Width) ? 1000 : this.Width);
            double h = this.ActualHeight > 0 ? this.ActualHeight : (double.IsNaN(this.Height) ? 360 : this.Height);

            this.Left = waLeft + (waWidth - w) / 2.0;
            this.Top = waTop + (waHeight - h);

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
                    SetProcessWorkingSetSize(System.Diagnostics.Process.GetCurrentProcess().Handle, -1, -1);
                }
            }
            catch { }
        }

        private void MainWindow_Loaded(object? sender, RoutedEventArgs e)
        {
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
            bool physShift = (GetAsyncKeyState(0x10) & 0x8000) != 0;
            bool physCtrl = (GetAsyncKeyState(0x11) & 0x8000) != 0;
            bool physAlt = (GetAsyncKeyState(0x12) & 0x8000) != 0;
            bool physWin = (GetAsyncKeyState(0x5B) & 0x8000) != 0;

            if (physShift)
            {
                if (!_lastPhysicalShiftDown)
                {
                    _physShiftDownTime = DateTime.Now;
                }
            }
            else
            {
                _physShiftDownTime = DateTime.MinValue;
            }

            _lastPhysicalShiftDown = physShift;

            _isCapsLockActive = (GetKeyState(0x14) & 0x0001) != 0;

            // 實體 Shift 視覺延遲邏輯：
            // 如果按住時間超過 500ms，才視為「啟動 Shift 版面預覽」
            // 這樣在短點按切換輸入法時，版面就不會跳動閃爍
            bool isPhysShiftLongPressed = physShift && (DateTime.Now - _physShiftDownTime).TotalMilliseconds > 200;
            _isShiftActive = isPhysShiftLongPressed || _virtualShiftToggle;
            _isCtrlActive = physCtrl || _virtualCtrlToggle;
            _isAltActive = physAlt || _virtualAltToggle;
            _isWinActive = physWin || _virtualWinToggle;
            _isNumLockActive = (GetKeyState(0x90) & 0x0001) != 0;

            foreach (var row in KeyRows)
            {
                foreach (var k in row)
                {
                    if (k.VkCode == MODE_KEY_CODE || k.VkCode == FN_KEY_CODE) continue;

                    bool isPhysicallyPressed = (GetAsyncKeyState(k.VkCode) & 0x8000) != 0;

                    if (k.VkCode == 0x10) isPhysicallyPressed = physShift;
                    else if (k.VkCode == 0x11) isPhysicallyPressed = physCtrl;
                    else if (k.VkCode == 0x12) isPhysicallyPressed = physAlt;

                    if (k.IsPressed != isPhysicallyPressed)
                    {
                        k.IsPressed = isPhysicallyPressed;
                    }
                }
            }

            _imeCheckCounter++;
            if (_imeCheckCounter >= 10)
            {
                _imeCheckCounter = 0;
                DetectImeStatus();
            }

            UpdateDisplay();
        }

        private void DetectImeStatus()
        {
            IntPtr foregroundHwnd = GetForegroundWindow();
            if (foregroundHwnd == IntPtr.Zero) return;

            uint threadId = GetWindowThreadProcessId(foregroundHwnd, out uint pid);
            string processName = "Unknown";
            try
            {
                var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                processName = proc.ProcessName;
            }
            catch { }

            GUITHREADINFO guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = Marshal.SizeOf(guiInfo);

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

        private void UpdateAppInfo(string processName, bool isZhuyin)
        {
            string status = isZhuyin ? "注音" : "英文";
            string clip = GetClipboardTextSafe();
            string newInfo = $"{processName} - {status}";
            if (!string.IsNullOrEmpty(clip))
            {
                newInfo += $" | 📋 {clip}";
            }

            if (_lastLoggedAppInfo != newInfo)
            {
                _lastLoggedAppInfo = newInfo;
                if (System.Windows.Application.Current != null && System.Windows.Application.Current.Dispatcher != null)
                {
                    System.Windows.Application.Current.Dispatcher.InvokeAsync(() => { CurrentAppInfo = newInfo; });
                }
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
                WINDOWPOS wp = Marshal.PtrToStructure<WINDOWPOS>(lParam);
                if ((wp.flags & SWP_NOMOVE) == 0 || (wp.flags & SWP_NOSIZE) == 0)
                {
                    double dpiX = 96.0, dpiY = 96.0;
                    var source = PresentationSource.FromVisual(this);
                    if (source?.CompositionTarget != null)
                    {
                        Matrix m = source.CompositionTarget.TransformToDevice;
                        dpiX = 96.0 * m.M11;
                        dpiY = 96.0 * m.M22;
                    }

                    int w = (wp.flags & SWP_NOSIZE) == 0 ? wp.cx : (int)(this.ActualWidth * dpiX / 96.0);
                    int h = (wp.flags & SWP_NOSIZE) == 0 ? wp.cy : (int)(this.ActualHeight * dpiY / 96.0);

                    if (w <= 0) w = (int)(1000 * dpiX / 96.0);
                    if (h <= 0) h = (int)(330 * dpiY / 96.0);

                    int curX = (wp.flags & SWP_NOMOVE) == 0 ? wp.x : (int)(this.Left * dpiX / 96.0);
                    int curY = (wp.flags & SWP_NOMOVE) == 0 ? wp.y : (int)(this.Top * dpiY / 96.0);

                    // Grab the monitor where this bounds rectangle would land
                    RECT rect = new RECT { left = curX, top = curY, right = curX + w, bottom = curY + h };
                    IntPtr hMonitor = MonitorFromRect(ref rect, MONITOR_DEFAULTTONEAREST);
                    if (hMonitor != IntPtr.Zero)
                    {
                        MONITORINFOEX mInfo = new MONITORINFOEX();
                        mInfo.cbSize = Marshal.SizeOf(mInfo);
                        if (GetMonitorInfo(hMonitor, ref mInfo))
                        {
                            var wa = mInfo.rcWork;

                            int newX = curX;
                            int newY = curY;

                            if (newX < wa.left) newX = wa.left;
                            if (newX + w > wa.right) newX = wa.right - w;

                            if (newY < wa.top) newY = wa.top;
                            // Make sure we check against wa.top again if wa.bottom - wa.top is actually smaller than the window (edge case)
                            if (newY + h > wa.bottom) newY = Math.Max(wa.top, wa.bottom - h);

                            if (newX != wp.x || newY != wp.y)
                            {
                                wp.x = newX;
                                wp.y = newY;
                                Marshal.StructureToPtr(wp, lParam, true);
                                handled = true;
                            }
                        }
                    }
                }
            }
            return IntPtr.Zero;
        }

        // Win32 API for native Monitor matching
        [DllImport("user32.dll")] private static extern IntPtr MonitorFromRect([In] ref RECT lprc, uint dwFlags);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);
        private const uint MONITOR_DEFAULTTONEAREST = 2;

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
            var transform = element.RenderTransform as ScaleTransform;
            if (transform == null || element.RenderTransformOrigin.X != 0.5)
            {
                transform = new ScaleTransform(1, 1);
                element.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
                element.RenderTransform = transform;
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
                                k.DynamicTextColor = _themeNumColor;
                            }
                            else
                            {
                                k.DisplayName = k.English;
                                k.DynamicTextColor = _themeTextColor;
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
                    if (k.VkCode == 0x90) { k.TextColor = _themeNumColor; k.IsActiveToggle = !_isNumLockActive; k.DisplayName = "Num"; k.DynamicTextColor = !_isNumLockActive ? _themeNumColor : _themeTextColor; }
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
            if (!File.Exists(_iniFilePath)) return;

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
                        this.Left = left;
                    }
                    else if (key == "Top" && double.TryParse(val, out double top))
                    {
                        this.Top = top;
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
                        this.Width = w;
                        double ratio = IsFullLayout ? 0.23 : 0.31;
                        double minH = Math.Max(w * ratio + 45, 120);
                        this.MinHeight = minH;
                        this.MaxHeight = minH;
                    }
                    else if (key == "Height" && double.TryParse(val, out double h))
                    {
                        this.Height = h;
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
                var lines = new List<string>
                {
                    $"Opacity={this.Opacity}",
                    $"Width={this.Width}",
                    $"Height={this.Height}",
                    $"Left={this.Left}",
                    $"Top={this.Top}",
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

        private void RestoreSize_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "確定要恢復預設配置嗎？\n您將會遺失所有自定義的按鍵位置與透明度設定。",
                "重置警告",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (result != MessageBoxResult.Yes) return;
            // 清除設定檔並恢復所有預設值
            this.Width = 1000;
            if (this.IsFullLayout)
            {
                ToggleFullLayout(null);
            }
            double ratio = 0.31;
            double minH = Math.Max(this.Width * ratio + 45, 120);
            this.MinHeight = minH;
            this.MaxHeight = minH;
            this.Height = minH;

            this.Opacity = 1.0;
            this.IsEditMode = false;

            _isPinned = true;
            OnPropertyChanged("PinIcon");
            this.Topmost = true; // 預設為 true

            this.IsDynamicLayout = true; // 預設為 true

            // 恢復顏色為系統預設
            DetectSystemTheme();

            if (OpacitySlider != null) OpacitySlider.Value = 1.0;
            MainRootGrid.Opacity = 1.0;

            if (File.Exists(_iniFilePath))
            {
                try { File.Delete(_iniFilePath); } catch { }
            }

            KeyRows.Clear();
            if (IsFullLayout) SetupFullKeyboard();
            else SetupKeyboard();

            KeyBoardItemsControl.ItemsSource = null;
            KeyBoardItemsControl.ItemsSource = KeyRows;

            // 讓視窗回到中央位置
            CenterWindow();
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 如果事件來源是觸控/手寫筆，或者是正在觸控拖曳狀態，則不執行 DragMove
            if (e.StylusDevice != null || _isTouchDraggingWindow) return;

            if (!_isResizing) this.DragMove();
        }

        private void TopTitleBar_TouchDown(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (IsEditMode) return;
            if (sender is UIElement el)
            {
                _isTouchDraggingWindow = true;
                // 記錄在視窗中的相對位置
                _touchDragStartPoint = e.GetTouchPoint(this).Position;
                el.CaptureTouch(e.TouchDevice);
                e.Handled = true; // 阻止向下轉換為滑鼠事件並呼叫 DragMove
            }
        }

        private void TopTitleBar_TouchMove(object sender, System.Windows.Input.TouchEventArgs e)
        {
            if (_isTouchDraggingWindow)
            {
                var pos = e.GetTouchPoint(this).Position;
                double dx = pos.X - _touchDragStartPoint.X;
                double dy = pos.Y - _touchDragStartPoint.Y;

                // 透過位移差調整視窗絕對位置 (此舉會將視窗瞬移，使相對 local 的指標位置回歸原點)
                this.Left += dx;
                this.Top += dy;
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
            }
        }
        private void Resize_Init(object sender, MouseButtonEventArgs e) { _isResizing = true; Mouse.Capture((UIElement)sender); }
        private void Resize_End(object sender, MouseButtonEventArgs e) { _isResizing = false; Mouse.Capture(null); }
        private void Resize_Move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (_isResizing)
            {
                System.Windows.Point p = e.GetPosition(this);
                if (p.X > 200) this.Width = p.X;

                double ratio = _isFullLayout ? 0.23 : 0.31;
                double minH = Math.Max(this.Width * ratio + 5, 80);
                this.MinHeight = minH;
                this.MaxHeight = minH;
                this.Height = minH;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _notifyIcon?.Dispose();
            _inputDetector?.Stop();
            if (!IsEditMode) SaveSettings(); // 確保關閉時儲存
            base.OnClosed(e);
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
        public string English { get { return _english; } set { _english = value; OnPropertyChanged("English"); OnPropertyChanged("EnglishDisp"); } }

        private string _englishUpper = "";
        public string EnglishUpper { get { return _englishUpper; } set { _englishUpper = value; OnPropertyChanged("EnglishUpper"); OnPropertyChanged("ShiftDisp"); } }

        private string _zhuyin = "";
        public string Zhuyin { get { return _zhuyin; } set { _zhuyin = value; OnPropertyChanged("Zhuyin"); OnPropertyChanged("ZhuyinDisp"); } }

        private string _fnText = "";
        public string FnText { get { return _fnText; } set { _fnText = value; OnPropertyChanged("FnText"); OnPropertyChanged("FnDisp"); } }

        private string _displayName = "";
        public string DisplayName { get { return _displayName; } set { _displayName = value; OnPropertyChanged("DisplayName"); } }

        private string _dynamicTextColor = "White";
        public string DynamicTextColor { get { return _dynamicTextColor; } set { _dynamicTextColor = value; OnPropertyChanged("DynamicTextColor"); } }

        public byte VkCode { get; set; }

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
                _height = value;
                OnPropertyChanged(nameof(Height));
            }
        }

        private double _leftMargin = 2;
        public double LeftMargin
        {
            get => _leftMargin;
            set
            {
                _leftMargin = value;
                OnPropertyChanged(nameof(LeftMargin));
                OnPropertyChanged(nameof(Margin));
            }
        }

        private double _topMargin = 2;
        public double TopMargin
        {
            get => _topMargin;
            set
            {
                _topMargin = value;
                OnPropertyChanged(nameof(TopMargin));
                OnPropertyChanged(nameof(Margin));
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
        public bool IsPressed { get { return _isPressed; } set { _isPressed = value; OnPropertyChanged("IsPressed"); OnPropertyChanged("Background"); } }

        private bool _isActiveToggle = false;
        public bool IsActiveToggle { get { return _isActiveToggle; } set { _isActiveToggle = value; OnPropertyChanged("IsActiveToggle"); OnPropertyChanged("Background"); } }

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