using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Interop;
using Inventor;
using UserControl = System.Windows.Controls.UserControl;

namespace MCG.Inventor.Ribbon
{
    /// <summary>
    /// Host chịu trách nhiệm nhúng WPF UserControl vào Inventor DockableWindow.
    /// Một instance ứng với 1 tool có palette.
    ///
    /// Đóng gói các quirk của Inventor đã gặp:
    ///   1. HWND=0 ngay cả trong kAfter — retry timer 200ms
    ///   2. Application.Current == null trong COM host — tạo instance thủ công
    ///   3. UpdateLayout() sau khi gán RootVisual
    ///   4. Sau 500ms resize lại vì DockableWindow chưa layout xong lần đầu
    ///   5. WM_SIZE của DockableWindow → resize HwndSource (qua NativeWindow subclass)
    ///   6. Inventor khôi phục Visible=true từ session trước → force hide 2s đầu
    /// </summary>
    internal class DockableWindowHost
    {
        private const string LOG_PREFIX = "[DockableWindowHost]";

        // ─── Win32 ────────────────────────────────────────────────────────────
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private const int WS_CHILD   = 0x40000000;
        private const int WS_VISIBLE = 0x10000000;

        // ─── Dependencies ─────────────────────────────────────────────────────
        private readonly global::Inventor.Application _app;
        private readonly IDockablePanelDescriptor     _descriptor;
        private readonly string                       _addinGuid;
        private readonly RibbonContext                _supportedContexts;

        // ─── State ────────────────────────────────────────────────────────────
        private DockableWindow                    _dockableWindow;
        private DockableWindowsEvents             _dockableWindowsEvents;
        private ApplicationEvents                 _appEvents;
        private HwndSource                        _hwndSource;
        private UserControl                       _content;
        private DockWindowSizer                   _sizer;
        private System.Windows.Forms.Timer        _embedRetryTimer;
        private System.Windows.Forms.Timer        _forceHideTimer;
        private bool                              _buttonClicked;
        private bool                              _wasVisibleBeforeSwitch;
        private ButtonDefinition                  _linkedButton;

        public string Id => _descriptor.Id;
        public UserControl Content => _content;

        public DockableWindowHost(
            global::Inventor.Application app,
            IDockablePanelDescriptor     descriptor,
            string                       addinGuid,
            RibbonContext                supportedContexts)
        {
            _app               = app        ?? throw new ArgumentNullException(nameof(app));
            _descriptor        = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
            _addinGuid         = addinGuid;
            _supportedContexts = supportedContexts;
        }

        public void LinkButton(ButtonDefinition button) => _linkedButton = button;

        public void Create()
        {
            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Create — bắt đầu.");

            var uiManager = _app.UserInterfaceManager;
            _dockableWindowsEvents = uiManager.DockableWindows.Events;

            try
            {
                _dockableWindow = uiManager.DockableWindows[_descriptor.Id];
                Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] DockableWindow đã tồn tại, tái sử dụng.");
                _dockableWindowsEvents.OnShow -= OnDockableWindowShow;
                _dockableWindowsEvents.OnShow += OnDockableWindowShow;
            }
            catch
            {
                Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] DockableWindow chưa tồn tại, tạo mới...");
                _dockableWindow = uiManager.DockableWindows.Add(
                    ClientId:     _addinGuid,
                    InternalName: _descriptor.Id,
                    Title:        _descriptor.Title
                );

                _dockableWindow.ShowTitleBar = true;
                _dockableWindow.Visible      = false;
                _dockableWindow.SetMinimumSize(_descriptor.MinWidth, _descriptor.MinHeight);
                _dockableWindow.DockingState = _descriptor.DefaultDockingState;

                _dockableWindowsEvents.OnShow += OnDockableWindowShow;
            }

            try { _dockableWindow.Visible = false; } catch { }
            ScheduleForceHideOnStartup();

            EmbedContent();

            try
            {
                _appEvents = _app.ApplicationEvents;
                _appEvents.OnActivateDocument += OnActivateDocument;
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI subscribe OnActivateDocument: {ex.Message}"); }

            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Create THÀNH CÔNG.");
        }

        public void Toggle()
        {
            if (_dockableWindow == null) return;
            _buttonClicked = true;
            StopForceHideTimer();

            bool newVisibility = !_dockableWindow.Visible;
            _dockableWindow.Visible = newVisibility;
            if (_linkedButton != null) _linkedButton.Pressed = newVisibility;

            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Toggle → Visible={newVisibility}");
        }

        public void Cleanup()
        {
            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Cleanup...");

            try
            {
                if (_appEvents != null)
                {
                    _appEvents.OnActivateDocument -= OnActivateDocument;
                    _appEvents = null;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI detach AppEvents: {ex.Message}"); }

            try
            {
                if (_dockableWindowsEvents != null)
                    _dockableWindowsEvents.OnShow -= OnDockableWindowShow;
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI detach dockable events: {ex.Message}"); }

            try
            {
                if (_embedRetryTimer != null)
                {
                    _embedRetryTimer.Stop();
                    _embedRetryTimer.Dispose();
                    _embedRetryTimer = null;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI dispose retry timer: {ex.Message}"); }

            try { StopForceHideTimer(); }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI dispose force hide timer: {ex.Message}"); }

            try
            {
                if (_sizer != null)
                {
                    _sizer.ReleaseHandle();
                    _sizer = null;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI release sizer: {ex.Message}"); }

            try
            {
                if (_hwndSource != null)
                {
                    _hwndSource.Dispose();
                    _hwndSource = null;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI dispose HwndSource: {ex.Message}"); }

            _content        = null;
            _dockableWindow = null;
            _linkedButton   = null;

            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Cleanup THÀNH CÔNG.");
        }

        // ─── Private: embed WPF content ───────────────────────────────────────

        private void EmbedContent()
        {
            if (System.Windows.Application.Current == null)
            {
                new System.Windows.Application
                {
                    ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                };
                Debug.WriteLine($"{LOG_PREFIX} WPF Application instance đã được tạo.");
            }

            if (_content == null)
            {
                AppDomain.CurrentDomain.AssemblyResolve += OnBamlAssemblyResolve;
                try
                {
                    _content = _descriptor.CreateContent();
                    Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Content đã tạo.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"{LOG_PREFIX} LỖI tạo content: {ex.Message}");
                    Debug.WriteLine($"{LOG_PREFIX} Stack:\n{ex.StackTrace}");
                    return;
                }
                finally
                {
                    AppDomain.CurrentDomain.AssemblyResolve -= OnBamlAssemblyResolve;
                }
            }

            int dockHwnd = _dockableWindow.HWND;
            if (dockHwnd == 0)
            {
                Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] HWND=0 — schedule retry 200ms.");
                ScheduleEmbedRetry();
                return;
            }
            StopEmbedRetry();

            IntPtr hwndParent = new IntPtr(dockHwnd);

            int w = _descriptor.MinWidth, h = _descriptor.MinHeight;
            if (GetClientRect(hwndParent, out RECT rect))
            {
                w = Math.Max(rect.Right - rect.Left, _descriptor.MinWidth);
                h = Math.Max(rect.Bottom - rect.Top, _descriptor.MinHeight);
            }

            if (_hwndSource != null)
            {
                _hwndSource.Dispose();
                _hwndSource = null;
            }

            var parameters = new HwndSourceParameters($"{_descriptor.Id}.WpfHost")
            {
                ParentWindow = hwndParent,
                WindowStyle  = WS_CHILD | WS_VISIBLE,
                Width        = w,
                Height       = h,
                PositionX    = 0,
                PositionY    = 0
            };

            _hwndSource = new HwndSource(parameters);
            _hwndSource.RootVisual = _content;
            _content.UpdateLayout();

            Debug.WriteLine($"{LOG_PREFIX} [{_descriptor.Id}] Embed THÀNH CÔNG: HWND={_hwndSource.Handle}, {w}x{h}px");

            var resizeTimer = new System.Windows.Forms.Timer { Interval = 500 };
            resizeTimer.Tick += (s, ev) =>
            {
                resizeTimer.Stop();
                resizeTimer.Dispose();
                try
                {
                    if (_hwndSource != null && _dockableWindow != null && _dockableWindow.HWND != 0)
                    {
                        if (GetClientRect(new IntPtr(_dockableWindow.HWND), out RECT r2))
                        {
                            int rw = Math.Max(r2.Right - r2.Left, _descriptor.MinWidth);
                            int rh = Math.Max(r2.Bottom - r2.Top, _descriptor.MinHeight);
                            MoveWindow(_hwndSource.Handle, 0, 0, rw, rh, true);
                            _content?.UpdateLayout();
                        }
                    }
                }
                catch { }
            };
            resizeTimer.Start();

            if (_sizer == null)
                _sizer = new DockWindowSizer(hwndParent, _hwndSource.Handle);
            else
                _sizer.UpdateChildHandle(_hwndSource.Handle);

            try { _descriptor.OnContentEmbedded(_content); }
            catch (Exception ex) { Debug.WriteLine($"{LOG_PREFIX} LỖI OnContentEmbedded: {ex.Message}"); }
        }

        private void ScheduleEmbedRetry()
        {
            if (_embedRetryTimer == null)
            {
                _embedRetryTimer = new System.Windows.Forms.Timer { Interval = 200 };
                _embedRetryTimer.Tick += (s, e) =>
                {
                    _embedRetryTimer.Stop();
                    EmbedContent();
                };
            }
            _embedRetryTimer.Stop();
            _embedRetryTimer.Start();
        }

        private void StopEmbedRetry()
        {
            _embedRetryTimer?.Stop();
        }

        // ─── Private: force hide 2s đầu ───────────────────────────────────────

        private void ScheduleForceHideOnStartup()
        {
            if (_forceHideTimer != null) return;

            _forceHideTimer = new System.Windows.Forms.Timer { Interval = 100 };
            int ticks = 0;
            _forceHideTimer.Tick += (s, e) =>
            {
                ticks++;
                if (_buttonClicked) { StopForceHideTimer(); return; }

                try
                {
                    if (_dockableWindow != null && _dockableWindow.Visible)
                    {
                        _dockableWindow.Visible = false;
                        if (_linkedButton != null) _linkedButton.Pressed = false;
                    }
                }
                catch { }

                if (ticks >= 20) StopForceHideTimer();
            };
            _forceHideTimer.Start();
        }

        private void StopForceHideTimer()
        {
            if (_forceHideTimer != null)
            {
                _forceHideTimer.Stop();
                _forceHideTimer.Dispose();
                _forceHideTimer = null;
            }
        }

        private static System.Reflection.Assembly OnBamlAssemblyResolve(object sender, ResolveEventArgs args)
        {
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (asm.FullName == args.Name) return asm;
            }
            return null;
        }

        // ─── Event handlers ───────────────────────────────────────────────────

        private void OnDockableWindowShow(
            DockableWindow dockableWindow,
            EventTimingEnum beforeOrAfter,
            NameValueMap context,
            out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (dockableWindow == null || dockableWindow.InternalName != _descriptor.Id) return;
            if (beforeOrAfter != EventTimingEnum.kAfter) return;

            if (_hwndSource == null || _content == null)
            {
                EmbedContent();
            }
            else
            {
                int dockHwnd = dockableWindow.HWND;
                if (dockHwnd != 0 && GetClientRect(new IntPtr(dockHwnd), out RECT r))
                {
                    int w = Math.Max(r.Right - r.Left, _descriptor.MinWidth);
                    int h = Math.Max(r.Bottom - r.Top, _descriptor.MinHeight);
                    MoveWindow(_hwndSource.Handle, 0, 0, w, h, true);
                    _content.UpdateLayout();
                }
            }
        }

        private void OnActivateDocument(
            _Document documentObject,
            EventTimingEnum beforeOrAfter,
            NameValueMap context,
            out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter != EventTimingEnum.kAfter) return;
            if (_dockableWindow == null) return;

            bool docMatchesContext = DocumentMatchesContext(documentObject);

            if (!docMatchesContext && _dockableWindow.Visible)
            {
                _wasVisibleBeforeSwitch = true;
                _dockableWindow.Visible = false;
                if (_linkedButton != null) _linkedButton.Pressed = false;
            }
            else if (docMatchesContext && _wasVisibleBeforeSwitch)
            {
                _wasVisibleBeforeSwitch = false;
                _dockableWindow.Visible = true;
                if (_linkedButton != null) _linkedButton.Pressed = true;
            }
        }

        private bool DocumentMatchesContext(_Document doc)
        {
            if (doc == null) return false;
            switch (doc.DocumentType)
            {
                case DocumentTypeEnum.kDrawingDocumentObject:
                    return _supportedContexts.HasFlag(RibbonContext.Drawing);
                case DocumentTypeEnum.kPartDocumentObject:
                    return _supportedContexts.HasFlag(RibbonContext.Part);
                case DocumentTypeEnum.kAssemblyDocumentObject:
                    return _supportedContexts.HasFlag(RibbonContext.Assembly);
                default:
                    return false;
            }
        }

        // ─── Nested: NativeWindow subclass để bắt WM_SIZE ─────────────────────

        private sealed class DockWindowSizer : NativeWindow
        {
            private const int WM_SIZE = 0x0005;

            [DllImport("user32.dll", SetLastError = true)]
            private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

            private IntPtr _hwndChild;

            public DockWindowSizer(IntPtr hwndParent, IntPtr hwndChild)
            {
                _hwndChild = hwndChild;
                AssignHandle(hwndParent);
            }

            public void UpdateChildHandle(IntPtr hwndChild) => _hwndChild = hwndChild;

            protected override void WndProc(ref Message m)
            {
                base.WndProc(ref m);
                if (m.Msg == WM_SIZE && _hwndChild != IntPtr.Zero)
                {
                    int width  = (int)(m.LParam.ToInt64() & 0xFFFF);
                    int height = (int)((m.LParam.ToInt64() >> 16) & 0xFFFF);
                    if (width > 0 && height > 0)
                        MoveWindow(_hwndChild, 0, 0, width, height, true);
                }
            }
        }
    }
}
