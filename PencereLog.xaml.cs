using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Text;
using Microsoft.Extensions.Logging;
using System.Windows.Controls;

namespace NetAi
{
    public partial class PencereLog : Window
    {

        private const int GWL_STYLE = -16;
        private const int WS_MINIMIZEBOX = 0x00020000;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int SC_MAXIMIZE = 0xF030;
        private const int SC_RESTORE = 0xF120;

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll")]
        private static extern bool DeleteMenu(IntPtr hMenu, int uPosition, int uFlags);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private static void SetWindowStyle(IntPtr hWnd, int nIndex, int newStyle)
        {
            if (Environment.Is64BitProcess)
                SetWindowLongPtr64(hWnd, nIndex, new IntPtr(newStyle));
            else
                SetWindowLong32(hWnd, nIndex, newStyle);
        }

        private Window ownerWindow;
        private Rect normalBounds;
        private bool isMaximized = false;
        private HwndSource hwndSource;
        private StringBuilder logHistory = new StringBuilder();
        private readonly ILogger<PencereLog> _logger;
        public System.Windows.Controls.TextBox LogTextBox { get; private set; }
        public PencereLog(Window owner)
        {
            InitializeComponent();
            ownerWindow = owner;
            Owner = owner;

            owner.LocationChanged += Owner_LocationChanged;
            owner.SizeChanged += Owner_SizeChanged;
            owner.StateChanged += Owner_StateChanged;
            owner.Closed += Owner_Closed;

            UpdatePosition();
            SourceInitialized += OnSourceInitialized;
            StateChanged += Window_StateChanged;
            LocationChanged += Window_LocationChanged;
        }

        public PencereLog(ILogger<PencereLog> logger, System.Windows.Controls.TextBox logTextBox)
        {
            InitializeComponent();

            LogTextBox = logTextBox;
            // Logger'ı constructor ile alıyoruz
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            // LogTextBox'ı burada tanımlıyoruz
            logTextBox = new System.Windows.Controls.TextBox();
            logTextBox.IsReadOnly = true;
            logTextBox.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            // TextBox'ı pencerede görüntülemek için bir yerleştirme
            var grid = new Grid();
            grid.Children.Add(logTextBox);
            this.Content = grid;

            // Logger kullanımı örneği
            _logger.LogInformation("PencereLog başlatıldı.");

        }

        public void Log(string message)
        {
            LogTextBox.AppendText(message + Environment.NewLine);
        }
        private void OnSourceInitialized(object sender, EventArgs e)
        {
            IntPtr handle = new WindowInteropHelper(this).Handle;
            hwndSource = PresentationSource.FromVisual(this) as HwndSource;

            if (hwndSource != null)
                hwndSource.AddHook(WndProc);

            DisableMinimizeButton(handle);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_SYSCOMMAND)
            {
                if (wParam.ToInt32() == SC_MAXIMIZE)
                {
                    normalBounds = new Rect(Left, Top, Width, Height);
                    isMaximized = true;
                }
                else if (wParam.ToInt32() == SC_RESTORE && isMaximized)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        WindowState = WindowState.Normal;
                        Left = normalBounds.Left;
                        Top = normalBounds.Top;
                        Width = normalBounds.Width;
                        Height = normalBounds.Height;
                    }));
                    isMaximized = false;
                    handled = true;
                    return IntPtr.Zero;
                }
            }
            return IntPtr.Zero;
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Maximized && !isMaximized)
            {
                normalBounds = new Rect(Left, Top, Width, Height);
                isMaximized = true;
            }
            else if (WindowState == WindowState.Normal && isMaximized)
            {
                Left = normalBounds.Left;
                Top = normalBounds.Top;
                Width = normalBounds.Width;
                Height = normalBounds.Height;
                isMaximized = false;
            }
        }

        private void Window_LocationChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Normal && !isMaximized)
            {
                normalBounds = new Rect(Left, Top, Width, Height);
            }
        }

        private void DisableMinimizeButton(IntPtr handle)
        {
            int currentStyle = GetWindowLong(handle, GWL_STYLE);
            int newStyle = currentStyle & ~WS_MINIMIZEBOX;
            SetWindowStyle(handle, GWL_STYLE, newStyle);
        }

        private void Owner_StateChanged(object sender, EventArgs e)
        {
            if (ownerWindow.WindowState == WindowState.Minimized)
                Hide();
            else
            {
                Show();
                UpdatePosition();
            }
        }

        private void Owner_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdatePosition();
        }

        private void Owner_LocationChanged(object sender, EventArgs e)
        {
            UpdatePosition();
        }

        private void Owner_Closed(object sender, EventArgs e)
        {
            logHistory.Clear();
            Close();
        }

        private void UpdatePosition()
        {
            if (ownerWindow == null || !ownerWindow.IsLoaded || ownerWindow.WindowState == WindowState.Minimized)
                return;

            try
            {
                var workingArea = ScreenHelper.GetWorkingArea(ownerWindow);
                double newLeft = ownerWindow.Left;
                double newTop = ownerWindow.Top + ownerWindow.ActualHeight + 5;

                if (newTop + ActualHeight > workingArea.Bottom)
                {
                    newTop = ownerWindow.Top - ActualHeight - 5;
                    if (newTop < workingArea.Top)
                        newTop = workingArea.Top;
                }

                if (ownerWindow.WindowState == WindowState.Maximized)
                {
                    newLeft = workingArea.Left;
                    newTop = workingArea.Top + ownerWindow.ActualHeight - 30;
                }

                Left = Math.Max(workingArea.Left, newLeft);
                Top = Math.Max(workingArea.Top, newTop);

                if (Left + Width > workingArea.Right)
                    Left = workingArea.Right - Width;

                Width = Math.Min(Width, ownerWindow.ActualWidth);
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Pencere konumu güncellenirken hata: {ex.Message}");
            }
        }

        public string TumLoglariAl() => logHistory.ToString();

        public void LoglariTemizle()
        {
            logHistory.Clear();
            Dispatcher.Invoke(() => LogTextBox?.Clear());
        }

        public static class ScreenHelper
        {
            public static Rect GetWorkingArea(Window window)
            {
                var handle = new WindowInteropHelper(window).Handle;
                var monitor = NativeMethods.MonitorFromWindow(handle, NativeMethods.MONITOR_DEFAULTTONEAREST);

                NativeMethods.MONITORINFO info = new NativeMethods.MONITORINFO();
                info.cbSize = Marshal.SizeOf(typeof(NativeMethods.MONITORINFO));

                NativeMethods.GetMonitorInfo(monitor, ref info);

                return new Rect(
                    info.rcWork.Left,
                    info.rcWork.Top,
                    info.rcWork.Right - info.rcWork.Left,
                    info.rcWork.Bottom - info.rcWork.Top);
            }

            private static class NativeMethods
            {
                public const int MONITOR_DEFAULTTONEAREST = 2;

                [DllImport("user32.dll")]
                public static extern IntPtr MonitorFromWindow(IntPtr hwnd, int flags);

                [DllImport("user32.dll")]
                public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

                [StructLayout(LayoutKind.Sequential)]
                public struct RECT
                {
                    public int Left, Top, Right, Bottom;
                }

                [StructLayout(LayoutKind.Sequential)]
                public struct MONITORINFO
                {
                    public int cbSize;
                    public RECT rcMonitor;
                    public RECT rcWork;
                    public int dwFlags;
                }
            }
        }
    }
}
