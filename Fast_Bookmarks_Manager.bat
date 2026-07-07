<# ::
    cls & @echo off & title Fast Bookmarks Manager

    for %%A in ("/?" "-?" "--?" "/help" "-help" "--help") do if /I "%~1"=="%%~A" goto :help
        
    if exist %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe   set "powershell=%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
    if exist %SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe  set "powershell=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    if not exist "%powershell%" set "powershell=powershell"

    set args=%*
    if defined args set args=%args:^=^^%
    if defined args set args=%args:<=^<%
    if defined args set args=%args:>=^>%
    if defined args set args=%args:&=^&%
    if defined args set args=%args:|=^|%
    if defined args set "args=%args:"=\"%"

    "%powershell%" -NoLogo -NoProfile -STA -Window Hidden -Command ^
        ^
        %= Create loading popup =% ^
        "$M=[Runtime.InteropServices.Marshal];" ^
        "$d=[AppDomain]::CurrentDomain.DefineDynamicAssembly(" ^
        "(New-Object Reflection.AssemblyName('W')),'Run').DefineDynamicModule('W');" ^
        "$t=$d.DefineType('A','Public,Class');" ^
        "$z=$t.DefinePInvokeMethod('CreateWindowExW','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([Int32],[String],[String],[Int32],[Int32],[Int32],[Int32],[Int32]," ^
        "[IntPtr],[IntPtr],[IntPtr],[IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('ShowWindow','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Bool])," ^
        "@([IntPtr],[Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('GetSystemMetrics','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Int32])," ^
        "@([Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('SendMessageW','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([IntPtr],[UInt32],[IntPtr],[IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('GetStockObject','gdi32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('InitCommonControlsEx','comctl32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Bool])," ^
        "@([IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$A=$t.CreateType();" ^
        "$sw=$A::GetSystemMetrics(0);$sh=$A::GetSystemMetrics(1);" ^
        "$hw=$A::CreateWindowExW(9,'#32770','Fast Bookmarks Manager',0x10C00000," ^
        "[int](($sw-440)/2),[int](($sh-130)/2),440,130," ^
        "[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        "$null=$A::ShowWindow($hw,5);" ^
        "$pc=$M::AllocHGlobal(8);$M::WriteInt32($pc,0,8);$M::WriteInt32($pc,4,0x20);" ^
        "$null=$A::InitCommonControlsEx($pc);$M::FreeHGlobal($pc);" ^
        "$ft=$A::GetStockObject(17);" ^
        "$hl=$A::CreateWindowExW(0,'Static','Initializing...',0x50000000," ^
        "20,15,390,20,$hw,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        "$null=$A::SendMessageW($hl,0x30,$ft,[IntPtr]::Zero);" ^
        "$hb=$A::CreateWindowExW(0,'msctls_progress32','',0x50000000," ^
        "20,42,390,24,$hw,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        ^
        %= PowerShell self-read, skipping batch part =% ^
        "$batFile='%~f0';$sb=[ScriptBlock]::Create([IO.File]::ReadAllText('%~f0'));& $sb @args" %args%

    for %%A in (%args%) do if /I "%%A"=="-autorestore" exit
    exit /b

    :help
    mode con: cols=120 lines=50
    echo.
    echo.
    echo    =============================================================================
    echo                              Fast Bookmarks Manager v1.0
    echo                                          ---
    echo                     Author : Leo Gillet - Freenitial on GitHub
    echo    =============================================================================
    echo.
    echo.
    echo    DESCRIPTION:
    echo       -----------
    echo       Fast Bookmarks Manager is a PowerShell tool with a graphical interface
    echo       for exporting and importing bookmarks-bar from Chrome / Edge / Firefox / OperaGX profiles.
    echo.
    echo       Key features:
    echo         - Export/Import bookmarks-bar as HTML or .url files
    echo         - Manage target folder for backup, and source folder for restoration
    echo         - Intuitive tree view for bookmarks selection
    echo         - Auto-restore for each default profile of each browser with the -autorestore parameter
    echo         - With Auto-restore you can add argument -source "C:\Folder_containing_url_files\" if not in script dir
    echo         - With Auto-restore you can add argument -source "C:\AnyPath\favs_file.html" if not in script dir
    echo         - When using -source argument, not followed by HTML file or folder path, HTML file in script dir will be prioritized
    echo.
    echo    OPTIONAL ARGUMENTS:
    echo       --------------------
    echo       1) -autorestore
    echo          - Enables automatic restoration mode without user prompts, for each default profile of each browser.
    echo.
    echo       2) "Full\Filepath\to\sourceFolder"
    echo          - Full path to the folder containing .url files for restoration. Script folder if not specified.
    echo.
    echo       3) "Full\Filepath\to\logfile.log"
    echo          - Full path for the log file. If not specified, TEMP folder will be used.
    echo.
    echo    USAGE:
    echo       ------
    echo       To launch Fast Bookmarks Manager normally: just open .bat file
    echo.
    echo       To launch with auto-restore and specify source and logfile paths:
    echo          start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager -autorestore "C:\Path\To\Source" "C:\Path\To\Logfile.log"
    echo.
    echo       Multi-line example:
    echo          start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager ^^
    echo                                   -autorestore ^^
    echo                                   -source "C:\Path\To\Source" ^^
    echo                                   -logfile "C:\Path\To\Logfile.log"
    echo.
    echo.
    echo    =============================================================================
    echo.
    pause>nul & exit /b
#>

param(
    [switch]$autorestore = $false,
    [string]$source,
    [string]$logfile
)

# ---- Remaining functions for Invoke-LoadingPump + updates ----
$t=$d.DefineType('E','Public,Class')
foreach($x in @(
    ,@('SetWindowTextW','user32.dll',([Bool]),@([IntPtr],[String]))
    ,@('DestroyWindow','user32.dll',([Bool]),@([IntPtr]))
    ,@('PeekMessageW','user32.dll',([Bool]),@([IntPtr],[IntPtr],[UInt32],[UInt32],[UInt32]))
    ,@('TranslateMessage','user32.dll',([Bool]),@([IntPtr]))
    ,@('DispatchMessageW','user32.dll',([IntPtr]),@([IntPtr]))
)){$z=$t.DefinePInvokeMethod($x[0],$x[1],'Public,Static,PinvokeImpl','Standard',$x[2],$x[3],'Winapi','Unicode');$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128)}
$E=$t.CreateType()

$mg=$M::AllocHGlobal(48)
function Invoke-LoadingPump{try{while($E::PeekMessageW($mg,[IntPtr]::Zero,0,0,1)){$null=$E::TranslateMessage($mg);$null=$E::DispatchMessageW($mg)}}catch{}}
function Update-LoadingPopup([int]$pct,[string]$s){$null=$A::SendMessageW($hb,0x402,[IntPtr]$pct,[IntPtr]::Zero);if($s){$null=$E::SetWindowTextW($hl,$s)};try{Invoke-LoadingPump}catch{}}
function Close-LoadingPopup{$null=$E::DestroyWindow($hw);try{Invoke-LoadingPump}catch{};$M::FreeHGlobal($mg)}
Update-LoadingPopup 5  "Loading..."

$OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# DPI aware + folder browser
Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
public class FolderSelectDialog {
    // credit this class 'Ben Philipp' https://stackoverflow.com/a/66823582
    private string _initialDirectory;
    private string _title;
    private string _message;
    private string _fileName = "";
    public string InitialDirectory {
        get { return string.IsNullOrEmpty(_initialDirectory) ? Environment.CurrentDirectory : _initialDirectory; }
        set { _initialDirectory = value; }
    }
    public string Title {
        get { return _title ?? "Select a folder"; }
        set { _title = value; }
    }
    public string Message {
        get { return _message ?? _title ?? "Select a folder"; }
        set { _message = value; }
    }
    public string FileName { get { return _fileName; } }

    public FolderSelectDialog(string defaultPath="MyComputer", string title="Select a folder", string message=""){
        InitialDirectory = defaultPath;
        Title = title;
        Message = message;
    }
    public bool Show() { return Show(IntPtr.Zero); }
    public bool Show(IntPtr? hWndOwnerNullable=null) {
        IntPtr hWndOwner = IntPtr.Zero;
        if(hWndOwnerNullable!=null)
            hWndOwner = (IntPtr)hWndOwnerNullable;
        if(Environment.OSVersion.Version.Major >= 6){
            try{
                var resulta = VistaDialog.Show(hWndOwner, InitialDirectory, Title, Message);
                _fileName = resulta.FileName;
                return resulta.Result;
            }
            catch(Exception){
                var resultb = ShowXpDialog(hWndOwner, InitialDirectory, Title, Message);
                _fileName = resultb.FileName;
                return resultb.Result;
            }
        }
        var result = ShowXpDialog(hWndOwner, InitialDirectory, Title, Message);
        _fileName = result.FileName;
        return result.Result;
    }
    private struct ShowDialogResult {
        public bool Result { get; set; }
        public string FileName { get; set; }
    }
    private static ShowDialogResult ShowXpDialog(IntPtr ownerHandle, string initialDirectory, string title, string message) {
        var folderBrowserDialog = new FolderBrowserDialog {
            Description = message,
            SelectedPath = initialDirectory,
            ShowNewFolderButton = true
        };
        var dialogResult = new ShowDialogResult();
        if (folderBrowserDialog.ShowDialog(new WindowWrapper(ownerHandle)) == DialogResult.OK) {
            dialogResult.Result = true;
            dialogResult.FileName = folderBrowserDialog.SelectedPath;
        }
        return dialogResult;
    }
    private static class VistaDialog {
        private const string c_foldersFilter = "Folders|\n";
        private const BindingFlags c_flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        private readonly static Assembly s_windowsFormsAssembly = typeof(FileDialog).Assembly;
        private readonly static Type s_iFileDialogType = s_windowsFormsAssembly.GetType("System.Windows.Forms.FileDialogNative+IFileDialog");
        private readonly static MethodInfo s_createVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("CreateVistaDialog", c_flags);
        private readonly static MethodInfo s_onBeforeVistaDialogMethodInfo = typeof(OpenFileDialog).GetMethod("OnBeforeVistaDialog", c_flags);
        private readonly static MethodInfo s_getOptionsMethodInfo = typeof(FileDialog).GetMethod("GetOptions", c_flags);
        private readonly static MethodInfo s_setOptionsMethodInfo = s_iFileDialogType.GetMethod("SetOptions", c_flags);
        private readonly static uint s_fosPickFoldersBitFlag = (uint) s_windowsFormsAssembly
            .GetType("System.Windows.Forms.FileDialogNative+FOS")
            .GetField("FOS_PICKFOLDERS")
            .GetValue(null);
        private readonly static ConstructorInfo s_vistaDialogEventsConstructorInfo = s_windowsFormsAssembly
            .GetType("System.Windows.Forms.FileDialog+VistaDialogEvents")
            .GetConstructor(c_flags, null, new[] { typeof(FileDialog) }, null);
        private readonly static MethodInfo s_adviseMethodInfo = s_iFileDialogType.GetMethod("Advise");
        private readonly static MethodInfo s_unAdviseMethodInfo = s_iFileDialogType.GetMethod("Unadvise");
        private readonly static MethodInfo s_showMethodInfo = s_iFileDialogType.GetMethod("Show");
        public static ShowDialogResult Show(IntPtr ownerHandle, string initialDirectory, string title, string description) {
            var openFileDialog = new OpenFileDialog {
                AddExtension = false,
                CheckFileExists = false,
                DereferenceLinks = true,
                Filter = c_foldersFilter,
                InitialDirectory = initialDirectory,
                Multiselect = false,
                Title = title
            };
            var iFileDialog = s_createVistaDialogMethodInfo.Invoke(openFileDialog, new object[] { });
            s_onBeforeVistaDialogMethodInfo.Invoke(openFileDialog, new[] { iFileDialog });
            s_setOptionsMethodInfo.Invoke(iFileDialog, new object[] { (uint) s_getOptionsMethodInfo.Invoke(openFileDialog, new object[] { }) | s_fosPickFoldersBitFlag });
            var adviseParametersWithOutputConnectionToken = new[] { s_vistaDialogEventsConstructorInfo.Invoke(new object[] { openFileDialog }), 0U };
            s_adviseMethodInfo.Invoke(iFileDialog, adviseParametersWithOutputConnectionToken);
            try {
                int retVal = (int) s_showMethodInfo.Invoke(iFileDialog, new object[] { ownerHandle });
                return new ShowDialogResult {
                    Result = retVal == 0,
                    FileName = openFileDialog.FileName
                };
            }
            finally {
                s_unAdviseMethodInfo.Invoke(iFileDialog, new[] { adviseParametersWithOutputConnectionToken[1] });
            }
        }
    }
    private class WindowWrapper : IWin32Window {
        private readonly IntPtr _handle;
        public WindowWrapper(IntPtr handle) { _handle = handle; }
        public IntPtr Handle { get { return _handle; } }
    }
    public string getPath(){
        if (Show()){
            return FileName;
        }
        return "";
    }
}
public static class DPIHelper {
    public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr dpiFlag);
    [DllImport("gdi32.dll")] static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
    public static float GetScaling() {
        using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) {
            IntPtr hdc = g.GetHdc();
            try {
                int dpi = GetDeviceCaps(hdc, 88);
                if (dpi > 0) { return (float)dpi / 96.0f; }
                return 1.0f;
            } finally { g.ReleaseHdc(hdc); }
        }
    }
}
public delegate void WndProcEventHandler(object sender, Message m);
public class CustomForm : Form {
    public event Action<float> DpiScaleChanged;
    private const int WM_DPICHANGED = 0x02E0;
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int left, top, right, bottom; }
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
    public CustomForm() {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        UpdateStyles();
    }
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_DPICHANGED) {
            int newDpi = m.WParam.ToInt32() & 0xFFFF;
            if (newDpi <= 0) newDpi = 96;
            float newScale = (float)newDpi / 96.0f;
            // Apply the OS-suggested window rectangle so the non-client frame is
            // resized to the new DPI, exactly as a fresh launch at that DPI would
            // be. The PS handler then only rescales the child controls, never the
            // window itself (a manual ClientSize left the frame inconsistent and
            // pushed the bottom-docked panel slightly out of the client area).
            RECT r = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
            // SWP_NOZORDER(0x4)|SWP_NOACTIVATE(0x10)|SWP_FRAMECHANGED(0x20)
            SetWindowPos(m.HWnd, IntPtr.Zero, r.left, r.top, r.right - r.left, r.bottom - r.top, 0x0034);
            if (DpiScaleChanged != null) DpiScaleChanged(newScale);
            m.Result = IntPtr.Zero;
            return;
        }
        base.WndProc(ref m);
    }
}
public static class DwmDark {
    [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr h, int a, ref int v, int s);
    [DllImport("uxtheme.dll", CharSet=CharSet.Unicode)] private static extern int SetWindowTheme(IntPtr h, string app, string idl);
    [DllImport("uxtheme.dll", EntryPoint="#135", SetLastError=true)] private static extern int SetPreferredAppMode(int mode);
    public static void Init() { try { SetPreferredAppMode(1); } catch {} }
    public static void Apply(IntPtr h, bool dark) {
        int v = dark ? 1 : 0;
        if (DwmSetWindowAttribute(h, 20, ref v, 4) != 0) { DwmSetWindowAttribute(h, 19, ref v, 4); }
    }
    public static void Scrollbars(IntPtr h, bool dark) {
        SetWindowTheme(h, dark ? "DarkMode_Explorer" : "Explorer", null);
    }
}
public static class DpiContext {
    public static float Scale = 1f;
    // Relative font ratio from the DPI at app start. Stays 1.0 while the app runs
    // at its startup DPI, then tracks new/startup on every live change. Under the
    // PowerShell host the control device contexts are frozen at the startup DPI,
    // so owner-drawn text (tab labels, group titles) multiplies its point size by
    // this to hold a constant physical size across live DPI changes.
    public static float FontScale = 1f;
    public static int S(int v) { return (int)(v * Scale); }
}
// A panel whose vertical scrollbar is ALWAYS visible (WS_VSCROLL + SIF_DISABLE-
// NOSCROLL) : disabled/greyed when nothing overflows (never hidden), functional
// when it does. Scrolling is MANUAL (no AutoScroll, which dynamically toggles the
// bar and makes it flicker/disappear) : it offsets its single content child.
public class ScrollPanel : Panel {
    private const int WS_VSCROLL = 0x00200000;
    private const int SB_VERT = 1, WM_VSCROLL = 0x0115;
    private const int SIF_RANGE = 0x1, SIF_PAGE = 0x2, SIF_POS = 0x4, SIF_DISABLENOSCROLL = 0x8, SIF_TRACKPOS = 0x10;
    private const uint ESB_ENABLE_BOTH = 0, ESB_DISABLE_BOTH = 3;
    [StructLayout(LayoutKind.Sequential)] private struct SCROLLINFO { public uint cbSize; public uint fMask; public int nMin, nMax, nPage, nPos, nTrackPos; }
    [DllImport("user32.dll")] private static extern int SetScrollInfo(IntPtr h, int bar, ref SCROLLINFO si, bool redraw);
    [DllImport("user32.dll")] private static extern bool GetScrollInfo(IntPtr h, int bar, ref SCROLLINFO si);
    [DllImport("user32.dll")] private static extern bool EnableScrollBar(IntPtr h, uint flags, uint arrows);
    [DllImport("user32.dll")] private static extern bool ShowScrollBar(IntPtr h, int bar, bool show);
    protected override CreateParams CreateParams {
        get { CreateParams cp = base.CreateParams; cp.Style |= WS_VSCROLL; return cp; }
    }
    private Control Content { get { return Controls.Count > 0 ? Controls[0] : null; } }
    private void UpdateScroll() {
        if (!IsHandleCreated) return;
        ShowScrollBar(Handle, SB_VERT, true);
        Control c = Content;
        int contentH = (c == null) ? 0 : c.Height;
        int clientH = ClientSize.Height;
        SCROLLINFO si = new SCROLLINFO(); si.cbSize = (uint)Marshal.SizeOf(typeof(SCROLLINFO));
        if (c != null && contentH > clientH) {
            int pos = -c.Top; if (pos > contentH - clientH) pos = contentH - clientH; if (pos < 0) pos = 0; c.Top = -pos;
            si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS; si.nMin = 0; si.nMax = contentH - 1; si.nPage = clientH; si.nPos = pos;
            SetScrollInfo(Handle, SB_VERT, ref si, true); EnableScrollBar(Handle, 1u, ESB_ENABLE_BOTH);
        } else {
            if (c != null && c.Top != 0) c.Top = 0;
            si.fMask = SIF_RANGE | SIF_PAGE | SIF_DISABLENOSCROLL; si.nMin = 0; si.nMax = 0; si.nPage = 1;
            SetScrollInfo(Handle, SB_VERT, ref si, true); EnableScrollBar(Handle, 1u, ESB_DISABLE_BOTH);
        }
    }
    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_VSCROLL) {
            int action = m.WParam.ToInt32() & 0xFFFF; Control c = Content;
            if (c != null) {
                int clientH = ClientSize.Height, contentH = c.Height, maxS = Math.Max(0, contentH - clientH), pos = -c.Top;
                SCROLLINFO si = new SCROLLINFO(); si.cbSize = (uint)Marshal.SizeOf(typeof(SCROLLINFO)); si.fMask = SIF_TRACKPOS; GetScrollInfo(Handle, SB_VERT, ref si);
                switch (action) { case 0: pos -= 24; break; case 1: pos += 24; break; case 2: pos -= clientH; break; case 3: pos += clientH; break; case 4: case 5: pos = si.nTrackPos; break; case 6: pos = 0; break; case 7: pos = maxS; break; }
                if (pos < 0) pos = 0; if (pos > maxS) pos = maxS; c.Top = -pos;
                SCROLLINFO s2 = new SCROLLINFO(); s2.cbSize = (uint)Marshal.SizeOf(typeof(SCROLLINFO)); s2.fMask = SIF_POS; s2.nPos = pos; SetScrollInfo(Handle, SB_VERT, ref s2, true);
            }
            return;
        }
        base.WndProc(ref m);
    }
    protected override void OnMouseWheel(MouseEventArgs e) {
        Control c = Content; int clientH = ClientSize.Height; int maxS = (c == null) ? 0 : Math.Max(0, c.Height - clientH);
        if (c == null || maxS <= 0) { base.OnMouseWheel(e); return; }
        int pos = -c.Top - e.Delta / 2; if (pos < 0) pos = 0; if (pos > maxS) pos = maxS; c.Top = -pos; UpdateScroll();
    }
    protected override void OnLayout(LayoutEventArgs e) { base.OnLayout(e); UpdateScroll(); }
    protected override void OnClientSizeChanged(EventArgs e) { base.OnClientSizeChanged(e); UpdateScroll(); }
}
public class DarkTabControl : TabControl {
    private int _hoverIdx = -1;
    public DarkTabControl() {
        SetStyle(ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.ResizeRedraw, true);
        DrawMode = TabDrawMode.OwnerDrawFixed;
        ApplyDpiScaling();
        SizeMode = TabSizeMode.Fixed;
    }
    protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); ApplyDpiScaling(); }
    public void ApplyDpiScaling() {
        float s = DpiContext.Scale; if (s <= 0f) s = 1f;
        Size ni = new Size((int)(96*s), (int)(30*s));
        Point np = new Point((int)(14*s), (int)(6*s));
        if (ItemSize != ni) ItemSize = ni;
        if (Padding != np) Padding = np;
    }
    protected override void OnMouseMove(MouseEventArgs e) {
        base.OnMouseMove(e); int n = -1;
        for (int i=0;i<TabCount;i++){ if (GetTabRect(i).Contains(e.Location)){ n=i; break; } }
        if (n != _hoverIdx){ _hoverIdx=n; Invalidate(); }
    }
    protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); _hoverIdx=-1; Invalidate(); }
    protected override void OnPaintBackground(PaintEventArgs e) {}
    protected override void OnPaint(PaintEventArgs e) {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.HighQuality;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        using (var b = new SolidBrush(Color.FromArgb(30,30,30))) g.FillRectangle(b, ClientRectangle);
        Rectangle sr = new Rectangle(0, 0, Width, ItemSize.Height + 4);
        using (var b = new SolidBrush(Color.FromArgb(26,26,28))) g.FillRectangle(b, sr);
        using (var p = new Pen(Color.FromArgb(38,38,42))) g.DrawLine(p, 0, sr.Bottom-1, Width, sr.Bottom-1);
        for (int i=0;i<TabCount;i++) {
            Rectangle r = GetTabRect(i); bool sel = (SelectedIndex==i); bool hov = (_hoverIdx==i);
            if (sel) {
                using (var b = new SolidBrush(Color.FromArgb(12,22,32))) g.FillRectangle(b, r);
                using (var b = new SolidBrush(Color.FromArgb(0,150,200))) g.FillRectangle(b, r.X, r.Bottom-2, r.Width, 2);
            } else if (hov) {
                using (var b = new SolidBrush(Color.FromArgb(22,28,36))) g.FillRectangle(b, r);
            }
            Color tc = sel ? Color.White : hov ? Color.FromArgb(210,215,220) : Color.FromArgb(130,135,140);
            using (var tf = new Font(Font.FontFamily, Font.Size * DpiContext.FontScale, Font.Style))
                TextRenderer.DrawText(g, TabPages[i].Text, tf, r, tc, TextFormatFlags.HorizontalCenter|TextFormatFlags.VerticalCenter|TextFormatFlags.NoPrefix);
        }
    }
}
public static class WinSqlite {
    const string DLL = "winsqlite3.dll";
    [DllImport(DLL, CharSet=CharSet.Unicode)] public static extern int sqlite3_open16(string f, out IntPtr db);
    [DllImport(DLL)] public static extern int sqlite3_close_v2(IntPtr db);
    [DllImport(DLL, CharSet=CharSet.Unicode)] public static extern int sqlite3_prepare16_v2(IntPtr db, string sql, int n, out IntPtr st, IntPtr tail);
    [DllImport(DLL)] public static extern int sqlite3_step(IntPtr st);
    [DllImport(DLL)] public static extern int sqlite3_finalize(IntPtr st);
    [DllImport(DLL)] public static extern IntPtr sqlite3_column_text16(IntPtr st, int c);
    [DllImport(DLL)] public static extern long sqlite3_column_int64(IntPtr st, int c);
    [DllImport(DLL)] public static extern int sqlite3_column_type(IntPtr st, int c);
    [DllImport(DLL, CharSet=CharSet.Unicode)] public static extern int sqlite3_bind_text16(IntPtr st, int i, string v, int n, IntPtr d);
    [DllImport(DLL)] public static extern int sqlite3_bind_int64(IntPtr st, int i, long v);
    [DllImport(DLL, CharSet=CharSet.Unicode)] public static extern IntPtr sqlite3_errmsg16(IntPtr db);
    public static string ColStr(IntPtr st, int c) { IntPtr p = sqlite3_column_text16(st, c); return p == IntPtr.Zero ? null : Marshal.PtrToStringUni(p); }
    public static long ColInt(IntPtr st, int c) { return sqlite3_column_int64(st, c); }
    public static string ErrMsg(IntPtr db) { IntPtr p = sqlite3_errmsg16(db); return p == IntPtr.Zero ? null : Marshal.PtrToStringUni(p); }
}
"@ -Language CSharp -ReferencedAssemblies @("System.Runtime.InteropServices", "System.Windows.Forms", "System.Drawing", "System.ComponentModel.Primitives")
[DPIHelper]::SetProcessDpiAwarenessContext([DPIHelper]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [DwmDark]::Init() } catch {}
$script:DPI_Factor = 1.0
try { $script:DPI_Factor = [DPIHelper]::GetScaling() } catch {}
if ($script:DPI_Factor -le 0) { $script:DPI_Factor = 1.0 }
try { [DpiContext]::Scale = [float]$script:DPI_Factor } catch {}
$script:StartupDpiFactor = $script:DPI_Factor
$script:SQLITE_TRANSIENT = New-Object IntPtr (-1)
$script:IsDark = $true

Update-LoadingPopup 20  "Loading..."

$script:ignoreAfterCheck = $false
$script:BackupSelectedBookmarks = @{}      # Dictionary: "Browser|Profile" => stored tree selection (with Checked states)
$script:RestoreSelectedUrls = @{}          # Dictionary: "Browser|Profile" => URL objects for restore
$script:CommonRestoreUrls = @{}            # Dictionary key "Common|Common" for common URL selection
$script:RestoreBaseFolder = ""             # Base for import path resolution (source folder, or virtual base for HTML)
$script:EnsureBookmarkBar = $false          # "Ensure bookmarks bar is always shown" checkbox (Restore tab)
$script:ExportErrors = @()
$script:sourcePath = (Get-Location).Path
$script:targetPath = (Get-Location).Path
$script:autorestore = $autorestore

function Test-PathAccess {
    param([string]$Path)
    try {
        if (Test-Path -LiteralPath $Path) {
            Get-ChildItem -LiteralPath $Path -ErrorAction Stop | Out-Null
            return $true
        }
    } catch { }
    return $false
}

function Test-FileWritable {
    param([string]$FilePath)
    try {
        $stream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Append, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $stream.Close()
        return $true
    } catch { return $false }
}

function Get-AvailableLogFileName {
    param([string]$FilePath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $ext = [System.IO.Path]::GetExtension($FilePath)
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    $suffix = 1
    $newPath = $FilePath
    while ((Test-Path -LiteralPath $newPath) -and (-not (Test-FileWritable -FilePath $newPath))) {
        $newPath = Join-Path $dir ("$base" + "_(" + $suffix + ")" + "$ext")
        $suffix++
    }
    return $newPath
}

# Logfile initialization
if (-not $logfile) { $logfile = Join-Path $env:TEMP "Fast_Bookmarks_Manager.log" } 
elseif (Test-Path -Path $logfile -PathType Container) { $logfile = Join-Path $logfile "Fast_Bookmarks_Manager.log" }

Write-Host "Detected logfile: $logfile"

# Check for invalid filename characters and replace with our usual title
$fileName = Split-Path -Path $logfile -Leaf
if ($fileName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
    Write-Host "Invalid filename '$fileName'. Replacing with 'Fast_Bookmarks_Manager.log'."
    $logfile = Join-Path (Split-Path -Path $logfile -Parent) "Fast_Bookmarks_Manager.log"
    Write-Host "New logfile path: $logfile"
}

$logDirectory = Split-Path -Path $logfile -Parent
Write-Host "Logfile directory: $logDirectory"

if (-not (Test-PathAccess -Path $logDirectory)) {
    Write-Host "Access denied or directory does not exist: $logDirectory"
    try {
        New-Item -ItemType Directory -Path $logDirectory -Force -ErrorAction Stop | Out-Null
        Write-Host "Directory created: $logDirectory"
    } catch {
        Write-Host "Failed to create directory: $logDirectory. Error: $($_.Exception.Message)"
        $logfile = Join-Path $env:TEMP "Fast_Bookmarks_Manager.log"
        Write-Host "Fallback logfile: $logfile"
    }
} else {
    Write-Host "Access confirmed for directory: $logDirectory"
}

$logDirectory = Split-Path -Path $logfile -Parent
if (-not (Test-PathAccess -Path $logDirectory)) {
    Write-Host "Access still denied after creation, fallback to TEMP."
    $logfile = Join-Path $env:TEMP "Fast_Bookmarks_Manager.log"
    Write-Host "New logfile: $logfile"
} else {
    Write-Host "Final access confirmed for: $logfile"
}

# Check if file exists and is locked or read-only; if so, get a new filename with recursive suffix
if (-not (Test-FileWritable -FilePath $logfile)) {
    Write-Host "File is locked or read-only. Searching for an available logfile."
    $logfile = Get-AvailableLogFileName -FilePath $logfile
    Write-Host "New logfile: $logfile"
}

function Log {
    param([Parameter(Mandatory = $true)][string]$message)
    Write-Host $message
    $message = "[$('{0:yyyy/MM/dd - HH:mm:ss}' -f (Get-Date))] - $message"
    try { Add-Content -Path $logfile -Value $message -ErrorAction Stop }
    catch { Write-Host "Error writing to logfile: $($_.Exception.Message)" }
}
Log "###############################"
Log "Logfile argument : $logfile"
Log "Source argument : $source"

# Restore-source precedence (see :help) : an explicit HTML file wins, then an
# explicit folder, otherwise an .html file sitting in the script directory is
# prioritised, and finally the script directory itself. $batFile is set by the
# .bat launcher ; fall back to the current location if run some other way.
$script:ScriptDir = if ($batFile -and (Test-Path -LiteralPath $batFile)) { Split-Path -Parent $batFile } else { (Get-Location).Path }
$script:sourcePath =
    if ($source -and ($source -match '\.html?$') -and (Test-Path -LiteralPath $source -PathType Leaf)) { $source }
    elseif ($source -and (Test-Path -LiteralPath $source -PathType Container)) { $source }
    else {
        $htmlHere = @(Get-ChildItem -LiteralPath $script:ScriptDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.html?$' })
        if ($htmlHere.Count -gt 0) { $htmlHere[0].FullName } else { $script:ScriptDir }
    }
Log "Resolved restore source : $script:sourcePath"

Update-LoadingPopup 30  "Loading..."

# ============================================================================
# BROWSER REGISTRY + PROFILE DISCOVERY
# ============================================================================
# Kind:
#   Chromium     - "User Data\<Profile>\Bookmarks" JSON, one folder per profile
#   ChromiumFlat - Opera : the profile folder itself holds "Bookmarks" (no sub-profiles)
#   Firefox      - places.sqlite, profiles listed in profiles.ini
$script:BrowserDefs = @(
    @{ Name="Chrome"; Kind="Chromium"; UserData="$env:LOCALAPPDATA\Google\Chrome\User Data"; Proc="chrome";
       Exe=@("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe", "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe") }
    @{ Name="Edge"; Kind="Chromium"; UserData="$env:LOCALAPPDATA\Microsoft\Edge\User Data"; Proc="msedge";
       Exe=@("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") }
    @{ Name="Brave"; Kind="Chromium"; UserData="$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"; Proc="brave";
       Exe=@("$env:ProgramFiles\BraveSoftware\Brave-Browser\Application\brave.exe", "${env:ProgramFiles(x86)}\BraveSoftware\Brave-Browser\Application\brave.exe", "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\Application\brave.exe") }
    @{ Name="Opera"; Kind="Chromium"; UserData="$env:APPDATA\Opera Software\Opera Stable"; Proc="opera";
       Exe=@("$env:LOCALAPPDATA\Programs\Opera\opera.exe", "$env:ProgramFiles\Opera\opera.exe", "${env:ProgramFiles(x86)}\Opera\opera.exe", "$env:LOCALAPPDATA\Programs\Opera\launcher.exe") }
    @{ Name="Opera GX"; Kind="OperaGX"; UserData="$env:APPDATA\Opera Software\Opera GX Stable"; Proc="opera";
       Exe=@("$env:LOCALAPPDATA\Programs\Opera GX\opera.exe", "$env:ProgramFiles\Opera GX\opera.exe", "${env:ProgramFiles(x86)}\Opera GX\opera.exe", "$env:LOCALAPPDATA\Programs\Opera GX\launcher.exe") }
    @{ Name="Firefox"; Kind="Firefox"; UserData="$env:APPDATA\Mozilla\Firefox"; Proc="firefox";
       Exe=@("$env:ProgramFiles\Mozilla Firefox\firefox.exe", "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe", "$env:LOCALAPPDATA\Mozilla Firefox\firefox.exe") }
)

function Resolve-BrowserExe($def) {
    foreach ($p in $def.Exe) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
    return $null
}

# Detect which browsers are present (user data exists).
function Get-DetectedBrowsers {
    $result = @()
    foreach ($def in $script:BrowserDefs) {
        $present = $false
        if ($def.Kind -eq "Firefox") { $present = (Test-Path -LiteralPath (Join-Path $def.UserData "profiles.ini")) }
        else { $present = (Test-Path -LiteralPath $def.UserData) }
        if ($present) {
            $d = $def.Clone()
            $d.ExePath = Resolve-BrowserExe $def
            $result += $d
            Log "Detected browser: $($def.Name) (exe: $($d.ExePath))"
        }
    }
    return $result
}

# Tolerant JSON parser : PS 5.1 ConvertFrom-Json throws on case-colliding keys
# (e.g. "name"/"Name" in the same object), which Edge's Local State contains.
# JavaScriptSerializer returns case-sensitive Dictionary<string,object> / object[].
function Get-JsonDict($text) {
    $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $ser.MaxJsonLength = [int]::MaxValue
    return $ser.DeserializeObject($text)
}

# Chromium profile order + display names are authoritative in "User Data\Local State"
# (profile.profiles_order + profile.info_cache.<dir>.name). Each profile's own
# Preferences carries a stale/default name, so it must NOT be used.
# Neither JSON parser handles every browser's Local State under PS 5.1 : native
# ConvertFrom-Json rejects Edge's (case-colliding keys), JavaScriptSerializer
# rejects Chrome's (recursion depth). Try native first, fall back to the other.
function Get-ChromiumLocalState($userData) {
    $result = @{ Order = @(); Names = @{} }
    $ls = Join-Path $userData "Local State"
    if (-not (Test-Path -LiteralPath $ls)) { return $result }
    $txt = Get-Content -LiteralPath $ls -Raw -Encoding UTF8
    $prof = $null; $useDict = $false
    try { $prof = ($txt | ConvertFrom-Json).profile }
    catch {
        try { $prof = (Get-JsonDict $txt)["profile"]; $useDict = $true } catch { return $result }
    }
    if (-not $prof) { return $result }
    if ($useDict) {
        if ($prof["profiles_order"]) { $result.Order = @($prof["profiles_order"]) }
        $ic = $prof["info_cache"]
        if ($ic) { foreach ($k in @($ic.Keys)) { $nm = $ic[$k]["name"]; if ($nm) { $result.Names[[string]$k] = [string]$nm } } }
    } else {
        if ($prof.profiles_order) { $result.Order = @($prof.profiles_order) }
        if ($prof.info_cache) { foreach ($p in $prof.info_cache.PSObject.Properties) { if ($p.Value.name) { $result.Names[$p.Name] = [string]$p.Value.name } } }
    }
    return $result
}

# Return profile objects for a detected browser :
#   @{ Browser; Kind; Id; Name; Source }  (Source = Bookmarks file or places.sqlite)
function Get-BrowserProfileList($def) {
    $list = @()
    switch ($def.Kind) {
        "Chromium" {
            $ls = Get-ChromiumLocalState $def.UserData
            $dirs = @()
            if ($ls.Order.Count -gt 0) { $dirs = @($ls.Order) }
            else {
                $dirs = @("Default")
                $dirs += Get-ChildItem -LiteralPath $def.UserData -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "^Profile \d+$" } | ForEach-Object { $_.Name }
            }
            foreach ($d in $dirs) {
                $bm = Join-Path $def.UserData (Join-Path $d "Bookmarks")
                $nm = if ($ls.Names.ContainsKey($d)) { $ls.Names[$d] } else { $d }
                $list += @{ Browser=$def.Name; Kind="Chromium"; Id=$d; Name=$nm; Source=$bm }
            }
            # Older Opera keeps a flat "<root>\Bookmarks" with no Default subfolder.
            $flat = Join-Path $def.UserData "Bookmarks"
            $hasProfileFile = $list | Where-Object { $_.Browser -eq $def.Name -and (Test-Path -LiteralPath $_.Source) }
            if (-not $hasProfileFile -and (Test-Path -LiteralPath $flat)) {
                $list = @($list | Where-Object { $_.Browser -ne $def.Name })
                $list += @{ Browser=$def.Name; Kind="Chromium"; Id="Default"; Name="Default"; Source=$flat }
            }
        }
        "OperaGX" {
            # Main profile.
            $ls = Get-ChromiumLocalState $def.UserData
            $mainName = if ($ls.Names.ContainsKey("Default") -and $ls.Names["Default"]) { $ls.Names["Default"] } else { "Default" }
            $list += @{ Browser=$def.Name; Kind="Chromium"; Id="Default"; Name=$mainName; Source=(Join-Path $def.UserData (Join-Path "Default" "Bookmarks")) }
            # Side profiles : _side_profiles\<hash>\<hash>_sideprofile.json (name/features),
            # bookmarks in _side_profiles\<hash>\Default\Bookmarks. Skip "roguelike" side
            # profiles (feature side-profile-clear-on-exit : ephemeral, cleared on exit).
            $sideRoot = Join-Path $def.UserData "_side_profiles"
            if (Test-Path -LiteralPath $sideRoot) {
                foreach ($sd in (Get-ChildItem -LiteralPath $sideRoot -Directory -ErrorAction SilentlyContinue)) {
                    $cfg = Get-ChildItem -LiteralPath $sd.FullName -Filter "*_sideprofile.json" -File -ErrorAction SilentlyContinue | Select-Object -First 1
                    if (-not $cfg) { continue }
                    try { $j = Get-Content -LiteralPath $cfg.FullName -Raw -Encoding UTF8 | ConvertFrom-Json } catch { continue }
                    $features = @(); if ($j.features) { $features = @($j.features) }
                    if ($features -contains "side-profile-clear-on-exit") { continue }
                    $nm = if ($j.name) { [string]$j.name } else { $sd.Name }
                    $list += @{ Browser=$def.Name; Kind="Chromium"; Id=("side:" + $sd.Name); Name=$nm; Source=(Join-Path $sd.FullName (Join-Path "Default" "Bookmarks")) }
                }
            }
        }
        "Firefox" {
            foreach ($prof in (Get-FirefoxProfiles $def.UserData)) {
                $list += @{ Browser=$def.Name; Kind="Firefox"; Id=$prof.Id; Name=$prof.Name; Source=$prof.Places }
            }
        }
    }
    Log "Found $($list.Count) profiles for $($def.Name)"
    return $list
}

# ============================================================================
# FIREFOX (places.sqlite via winsqlite3.dll)
# ============================================================================
function Get-FirefoxProfiles($ffRoot) {
    $ini = Join-Path $ffRoot "profiles.ini"
    if (-not (Test-Path -LiteralPath $ini)) { return @() }
    $text = Get-Content -LiteralPath $ini -Raw -Encoding UTF8
    $profiles = @()
    foreach ($m in [regex]::Matches($text, '(?ms)^\[Profile\d+\](.*?)(?=^\[|\z)')) {
        $sec = $m.Groups[1].Value
        $path = [regex]::Match($sec, '(?m)^Path=(.+?)\s*$').Groups[1].Value
        if (-not $path) { continue }
        $name = [regex]::Match($sec, '(?m)^Name=(.+?)\s*$').Groups[1].Value
        $isRel = [regex]::Match($sec, '(?m)^IsRelative=(\d)').Groups[1].Value
        $dir = if ($isRel -eq "0") { $path } else { Join-Path $ffRoot ($path -replace '/', '\') }
        $places = Join-Path $dir "places.sqlite"
        if (Test-Path -LiteralPath $places) {
            $profName = if ($name) { $name } else { Split-Path $dir -Leaf }
            $profiles += @{ Id=$path; Name=$profName; Places=$places }
        }
    }
    return $profiles
}

# Firefox url_hash : (HashScheme & 0xFFFF) << 32 | HashFull  (golden-ratio rotate hash).
function Get-FirefoxUrlHash([string]$url) {
    $mask = [uint64]4294967295; $golden = [uint64]2654435769
    $hashBytes = {
        param($bytes, $len)
        $h = [uint64]0
        for ($i = 0; $i -lt $len; $i++) {
            $rot = ((($h -shl 5) -band $mask) -bor ($h -shr 27)) -band $mask
            $mixed = ($rot -bxor ([uint64]$bytes[$i])) -band $mask
            $h = ($golden * $mixed) -band $mask
        }
        return [uint64]$h
    }
    $full = [System.Text.Encoding]::UTF8.GetBytes($url)
    $fullHash = & $hashBytes $full ([Math]::Min($full.Length, 1500))
    $ci = $url.IndexOf(':')
    $scheme = if ($ci -ge 0) { $url.Substring(0, $ci) } else { $url }
    $sb = [System.Text.Encoding]::UTF8.GetBytes($scheme)
    $schemeHash = (& $hashBytes $sb $sb.Length) -band 65535
    return [long](([uint64]$schemeHash -shl 32) -bor $fullHash)
}

function Open-SqliteDb([string]$path) {
    $db = [IntPtr]::Zero
    if ([WinSqlite]::sqlite3_open16($path, [ref]$db) -ne 0) { throw "sqlite open failed: $path" }
    return $db
}
function Invoke-Sqlite([IntPtr]$db, [string]$sql, [object[]]$binds = @()) {
    $st = [IntPtr]::Zero
    if ([WinSqlite]::sqlite3_prepare16_v2($db, $sql, -1, [ref]$st, [IntPtr]::Zero) -ne 0) { throw ("sqlite prepare: " + [WinSqlite]::ErrMsg($db)) }
    try {
        for ($i = 0; $i -lt $binds.Count; $i++) {
            $v = $binds[$i]
            if ($v -is [string]) { [void][WinSqlite]::sqlite3_bind_text16($st, $i + 1, $v, -1, $script:SQLITE_TRANSIENT) }
            else { [void][WinSqlite]::sqlite3_bind_int64($st, $i + 1, [long]$v) }
        }
        $rc = [WinSqlite]::sqlite3_step($st)
        if ($rc -ne 100 -and $rc -ne 101) { throw ("sqlite step rc=$rc : " + [WinSqlite]::ErrMsg($db)) }
    } finally { [void][WinSqlite]::sqlite3_finalize($st) }
}
function Get-SqliteRows([IntPtr]$db, [string]$sql, [int]$cols, [object[]]$binds = @()) {
    $st = [IntPtr]::Zero
    if ([WinSqlite]::sqlite3_prepare16_v2($db, $sql, -1, [ref]$st, [IntPtr]::Zero) -ne 0) { throw ("sqlite prepare: " + [WinSqlite]::ErrMsg($db)) }
    $rows = New-Object System.Collections.Generic.List[object]
    try {
        for ($i = 0; $i -lt $binds.Count; $i++) {
            $v = $binds[$i]
            if ($v -is [string]) { [void][WinSqlite]::sqlite3_bind_text16($st, $i + 1, $v, -1, $script:SQLITE_TRANSIENT) }
            else { [void][WinSqlite]::sqlite3_bind_int64($st, $i + 1, [long]$v) }
        }
        while ([WinSqlite]::sqlite3_step($st) -eq 100) {
            $row = @()
            for ($c = 0; $c -lt $cols; $c++) {
                if ([WinSqlite]::sqlite3_column_type($st, $c) -eq 1) { $row += [WinSqlite]::ColInt($st, $c) }
                else { $row += [WinSqlite]::ColStr($st, $c) }
            }
            [void]$rows.Add($row)
        }
    } finally { [void][WinSqlite]::sqlite3_finalize($st) }
    return ,$rows.ToArray()
}

# Copy a (possibly locked) DB + wal/shm to a temp dir so it can be opened read-only.
function Copy-SqliteForRead([string]$dbPath) {
    $tmp = Join-Path $env:TEMP ("fbm_ff_" + [IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    Copy-Item -LiteralPath $dbPath (Join-Path $tmp "places.sqlite") -Force
    foreach ($e in "-wal", "-shm") { if (Test-Path -LiteralPath ($dbPath + $e)) { Copy-Item -LiteralPath ($dbPath + $e) (Join-Path $tmp ("places.sqlite" + $e)) -Force } }
    return $tmp
}

# Read the Firefox bookmarks toolbar into the unified node model.
function Read-FirefoxBookmarks([string]$placesPath) {
    $tmp = Copy-SqliteForRead $placesPath
    $db = Open-SqliteDb (Join-Path $tmp "places.sqlite")
    try {
        $root = Get-SqliteRows $db "SELECT id FROM moz_bookmarks WHERE guid='toolbar_____'" 1
        if ($root.Count -eq 0) { return @() }
        $rootId = [long]$root[0][0]
        function Read-FfFolder($db, $parentId) {
            $items = Get-SqliteRows $db "SELECT b.type, b.title, p.url FROM moz_bookmarks b LEFT JOIN moz_places p ON b.fk=p.id WHERE b.parent=? ORDER BY b.position" 3 @($parentId)
            $out = @()
            $idx = 0
            $children = Get-SqliteRows $db "SELECT b.id, b.type, b.title, p.url FROM moz_bookmarks b LEFT JOIN moz_places p ON b.fk=p.id WHERE b.parent=? ORDER BY b.position" 4 @($parentId)
            foreach ($c in $children) {
                $id = [long]$c[0]; $type = [long]$c[1]; $title = [string]$c[2]; $url = [string]$c[3]
                if ($type -eq 1 -and $url) {
                    $nm = if ([string]::IsNullOrWhiteSpace($title)) { $url } else { $title }
                    $out += @{ Text=$nm; Tag=$url; Type="Bookmark"; Checked=$true }
                } elseif ($type -eq 2) {
                    $nm = if ([string]::IsNullOrWhiteSpace($title)) { "Folder" } else { $title }
                    $kids = Read-FfFolder $db $id
                    if ($kids.Count -gt 0) { $out += @{ Text=$nm; Type="Folder"; Children=$kids; Checked=$true } }
                }
            }
            return $out
        }
        return Read-FfFolder $db $rootId
    } finally {
        [void][WinSqlite]::sqlite3_close_v2($db)
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Insert url objects under the Firefox toolbar, recreating relative-path folders.
# Firefox must be closed. The DB is backed up first.
function Import-FirefoxUrls([string]$placesPath, [array]$urlObjects, [string]$baseFolder) {
    if ($urlObjects.Count -eq 0) { return }
    Copy-Item -LiteralPath $placesPath ($placesPath + ".fbm-backup") -Force -ErrorAction SilentlyContinue
    $db = Open-SqliteDb $placesPath
    try {
        $now = [long]((([DateTime]::UtcNow) - (New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc))).Ticks / 10)
        $rnd = New-Object Random
        function New-FfGuid($rnd) {
            $rb = New-Object byte[] 9; $rnd.NextBytes($rb)
            return [Convert]::ToBase64String($rb).Replace('+', '-').Replace('/', '_').Substring(0, 12)
        }
        $toolbarId = [long](Get-SqliteRows $db "SELECT id FROM moz_bookmarks WHERE guid='toolbar_____'" 1)[0][0]
        Invoke-Sqlite $db "BEGIN"
        try {
            foreach ($obj in $urlObjects) {
                # Recreate the relative folder path under the toolbar.
                $parentId = $toolbarId
                $fileDir = [System.IO.Path]::GetDirectoryName($obj.Path)
                $rel = ""
                if ($fileDir -and $baseFolder -and ($fileDir.TrimEnd('\') -ne $baseFolder.TrimEnd('\'))) { try { $rel = Get-RelativePath $fileDir $baseFolder } catch {} }
                if ($rel -ne "" -and $rel -ne ".") {
                    foreach ($folder in ($rel -split "\\")) {
                        if ($folder -eq "") { continue }
                        $existing = Get-SqliteRows $db "SELECT id FROM moz_bookmarks WHERE parent=? AND type=2 AND title=?" 1 @($parentId, $folder)
                        if ($existing.Count -gt 0) { $parentId = [long]$existing[0][0]; continue }
                        $pos = [long](Get-SqliteRows $db "SELECT COALESCE(MAX(position)+1,0) FROM moz_bookmarks WHERE parent=?" 1 @($parentId))[0][0]
                        Invoke-Sqlite $db "INSERT INTO moz_bookmarks (type,parent,position,title,dateAdded,lastModified,guid,syncStatus,syncChangeCounter) VALUES (2,?,?,?,?,?,?,1,1)" @($parentId, $pos, $folder, $now, $now, (New-FfGuid $rnd))
                        $parentId = [long](Get-SqliteRows $db "SELECT last_insert_rowid()" 1)[0][0]
                    }
                }
                # Skip if the url already exists under this parent.
                $dup = Get-SqliteRows $db "SELECT b.id FROM moz_bookmarks b JOIN moz_places p ON b.fk=p.id WHERE b.parent=? AND p.url=?" 1 @($parentId, $obj.URL)
                if ($dup.Count -gt 0) { continue }
                # Reuse an existing place row, else create one.
                $ex = Get-SqliteRows $db "SELECT id FROM moz_places WHERE url=?" 1 @($obj.URL)
                if ($ex.Count -gt 0) {
                    $placeId = [long]$ex[0][0]
                    Invoke-Sqlite $db "UPDATE moz_places SET foreign_count=foreign_count+1 WHERE id=?" @($placeId)
                }
                else {
                    $h2 = [regex]::Match($obj.URL, '://([^/:]+)').Groups[1].Value
                    $rev = ""
                    if ($h2) { $ra = $h2.ToCharArray(); [Array]::Reverse($ra); $rev = (-join $ra) + '.' }
                    $hash = Get-FirefoxUrlHash $obj.URL
                    Invoke-Sqlite $db "INSERT INTO moz_places (url,title,rev_host,hidden,typed,frecency,guid,foreign_count,url_hash) VALUES (?,?,?,0,0,-1,?,1,?)" @($obj.URL, $obj.Name, $rev, (New-FfGuid $rnd), $hash)
                    $placeId = [long](Get-SqliteRows $db "SELECT last_insert_rowid()" 1)[0][0]
                }
                $pos = [long](Get-SqliteRows $db "SELECT COALESCE(MAX(position)+1,0) FROM moz_bookmarks WHERE parent=?" 1 @($parentId))[0][0]
                Invoke-Sqlite $db "INSERT INTO moz_bookmarks (type,fk,parent,position,title,dateAdded,lastModified,guid,syncStatus,syncChangeCounter) VALUES (1,?,?,?,?,?,?,?,1,1)" @($placeId, $parentId, $pos, $obj.Name, $now, $now, (New-FfGuid $rnd))
            }
            Invoke-Sqlite $db "UPDATE moz_bookmarks SET lastModified=?, syncChangeCounter=syncChangeCounter+1 WHERE id=?" @($now, $toolbarId)
            Invoke-Sqlite $db "COMMIT"
        } catch { Invoke-Sqlite $db "ROLLBACK"; throw }
    } finally { [void][WinSqlite]::sqlite3_close_v2($db) }
}

# ============================================================================
# UNIFIED BOOKMARK READ / COUNT / IMPORT (dispatch by profile kind)
# ============================================================================
function Read-ChromiumBookmarks([string]$bmFile) {
    if (-not (Test-Path -LiteralPath $bmFile)) { return @() }
    $json = Get-Content -LiteralPath $bmFile -Raw -Encoding UTF8 | ConvertFrom-Json
    function Read-ChromeNodes($nodes) {
        $r = @()
        foreach ($n in $nodes) {
            if ($n.type -eq 'url') {
                $nm = if ([string]::IsNullOrWhiteSpace($n.name)) { $n.url } else { [System.Web.HttpUtility]::HtmlDecode($n.name) }
                $r += @{ Text=$nm; Tag=$n.url; Type="Bookmark"; Checked=$true }
            } elseif ($n.children) {
                $nm = if ([string]::IsNullOrWhiteSpace($n.name)) { "Folder" } else { [System.Web.HttpUtility]::HtmlDecode($n.name) }
                $c = Read-ChromeNodes $n.children
                if ($c.Count -gt 0) { $r += @{ Text=$nm; Type="Folder"; Children=$c; Checked=$true } }
            }
        }
        return $r
    }
    return Read-ChromeNodes $json.roots.bookmark_bar.children
}

# Unified read : returns the node model for any profile object.
function Read-ProfileBookmarks($profile) {
    if ($profile.Kind -eq "Firefox") { return Read-FirefoxBookmarks $profile.Source }
    return Read-ChromiumBookmarks $profile.Source
}

function Get-NodeUrlCount($nodes) {
    $c = 0
    foreach ($n in $nodes) {
        if ($n.Type -eq "Bookmark") { $c++ }
        elseif ($n.Type -eq "Folder") { $c += Get-NodeUrlCount $n.Children }
    }
    return $c
}

function Test-ProfileHasBookmarks($profile) {
    if ($profile.Kind -eq "Firefox") { return (Test-Path -LiteralPath $profile.Source) }
    return (Test-Path -LiteralPath $profile.Source)
}

# ============================================================================
# THEME + SCALING
# ============================================================================
$script:Theme = @{
    Back        = [System.Drawing.Color]::FromArgb(45, 45, 48)
    Fore        = [System.Drawing.Color]::White
    ForeDim     = [System.Drawing.Color]::FromArgb(160, 160, 160)
    ControlBack = [System.Drawing.Color]::FromArgb(55, 55, 55)
    Border      = [System.Drawing.Color]::FromArgb(80, 80, 80)
    GroupBorder = [System.Drawing.Color]::FromArgb(50, 90, 120)
    BtnBack     = [System.Drawing.Color]::FromArgb(60, 60, 60)
    TabBack     = [System.Drawing.Color]::FromArgb(32, 32, 32)
    TabActive   = [System.Drawing.Color]::FromArgb(0, 76, 127)
    TreeBack    = [System.Drawing.Color]::FromArgb(30, 30, 30)
    Accent      = [System.Drawing.Color]::FromArgb(0, 120, 212)
}
$script:GroupPaintHandler = {
    param($s, $e)
    $g = $e.Graphics
    $g.Clear($s.BackColor)
    # The title self-scales by FontScale (the container font stays at the base
    # size so leaf controls inheriting it are not double-scaled by the tree walk).
    $fs = [DpiContext]::FontScale; if ($fs -le 0) { $fs = 1 }
    $tf = New-Object System.Drawing.Font($s.Font.FontFamily, ([float]$s.Font.Size * $fs), [System.Drawing.FontStyle]::Bold)
    $pen = New-Object System.Drawing.Pen($script:Theme.GroupBorder)
    $ts = $g.MeasureString($s.Text, $tf)
    $half = [int]($ts.Height / 2)
    $rect = New-Object System.Drawing.Rectangle(0, $half, ($s.Width - 1), ($s.Height - $half - 1))
    $g.DrawLine($pen, $rect.X, $rect.Y, 8, $rect.Y)
    $g.DrawLine($pen, (8 + [int]$ts.Width + 2), $rect.Y, $rect.Right, $rect.Y)
    $g.DrawLine($pen, $rect.X, $rect.Y, $rect.X, $rect.Bottom)
    $g.DrawLine($pen, $rect.X, $rect.Bottom, $rect.Right, $rect.Bottom)
    $g.DrawLine($pen, $rect.Right, $rect.Y, $rect.Right, $rect.Bottom)
    $pen.Dispose()
    $br = New-Object System.Drawing.SolidBrush($s.ForeColor)
    $g.DrawString($s.Text, $tf, $br, 10, 0)
    $br.Dispose()
    $tf.Dispose()
}
function Set-DarkTitleBar($form) {
    try { if ($script:IsDark -and $form.IsHandleCreated) { [DwmDark]::Apply($form.Handle, $true) } } catch {}
}
# Apply the dark palette across a control subtree.
function Apply-DarkTheme($root) {
    if (-not $script:IsDark) { return }
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $c = $stack.Pop()
        foreach ($ch in $c.Controls) { $stack.Push($ch) }
        try {
            if ($c -is [System.Windows.Forms.TextBox]) {
                $c.BackColor = $script:Theme.ControlBack; $c.ForeColor = $script:Theme.Fore
                $c.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            } elseif ($c -is [System.Windows.Forms.Button]) {
                $c.BackColor = $script:Theme.BtnBack; $c.ForeColor = $script:Theme.Fore
                $c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $c.FlatAppearance.BorderColor = $script:Theme.Border
            } elseif ($c -is [System.Windows.Forms.TreeView]) {
                $c.BackColor = $script:Theme.TreeBack; $c.ForeColor = $script:Theme.Fore
                $c.LineColor = $script:Theme.Border
            } elseif ($c -is [DarkTabControl]) {
                # self-painted
            } elseif ($c -is [System.Windows.Forms.GroupBox]) {
                $c.ForeColor = $script:Theme.Fore; $c.BackColor = $script:Theme.Back
            } else {
                $c.BackColor = $script:Theme.Back; $c.ForeColor = $script:Theme.Fore
            }
        } catch {}
    }
}
# Dark-theme the native scrollbars of scrollable controls (SetWindowTheme).
function Set-DarkScrollbars($root) {
    if (-not $script:IsDark) { return }
    $types = @('FlowLayoutPanel', 'Panel', 'TreeView', 'ListBox', 'ScrollPanel')
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $c = $stack.Pop()
        foreach ($ch in $c.Controls) { $stack.Push($ch) }
        try { if ($c.IsHandleCreated -and ($types -contains $c.GetType().Name)) { [DwmDark]::Scrollbars($c.Handle, $true) } } catch {}
    }
}
# Manual DPI scaling of a control subtree. sizeFactor scales bounds ; fontFactor
# scales fonts (1.0 = leave them : a freshly-created control renders at the
# monitor DPI already, only runtime re-scaling of existing controls needs a
# font factor to compensate their frozen device context).
function Scale-ControlTree($root, $sizeFactor, $fontFactor) {
    $cache = @{}
    $getFont = {
        param($f, $r)
        $sz = [float]$f.Size * [float]$r
        $k = '{0}|{1}|{2}' -f $f.FontFamily.Name, $sz, [int]$f.Style
        if ($cache.ContainsKey($k)) { return $cache[$k] }
        $nf = New-Object System.Drawing.Font($f.FontFamily, $sz, $f.Style)
        $cache[$k] = $nf; return $nf
    }
    $doFont = ([Math]::Abs([float]$fontFactor - 1.0) -ge 0.001)
    $doSize = ([Math]::Abs([float]$sizeFactor - 1.0) -ge 0.001)
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $c = $stack.Pop()
        foreach ($ch in $c.Controls) { $stack.Push($ch) }
        if ($c -eq $root) { continue }
        try {
            # Leaf controls only. Scaling a CONTAINER font (GroupBox, TabControl,
            # Panel) would compound : every descendant that inherits that font
            # gets multiplied here too, then again when the walk reaches the leaf
            # itself -> ratio^2. Containers keep the base font ; their owner-drawn
            # text (tab labels, group titles) self-scales via DpiContext.FontScale.
            $fontBearing = ($c -is [System.Windows.Forms.Label]) -or ($c -is [System.Windows.Forms.Button]) -or ($c -is [System.Windows.Forms.TextBox]) -or ($c -is [System.Windows.Forms.CheckBox]) -or ($c -is [System.Windows.Forms.RadioButton]) -or ($c -is [System.Windows.Forms.LinkLabel]) -or ($c -is [System.Windows.Forms.TreeView])
            if ($doFont -and $fontBearing -and $null -ne $c.Font) { $c.Font = & $getFont $c.Font $fontFactor }
        } catch {}
        try {
            if ($doSize) {
                $dock = $c.Dock
                if ($dock -eq [System.Windows.Forms.DockStyle]::None) {
                    if (-not $c.AutoSize) { $c.Size = New-Object System.Drawing.Size([int]($c.Width * $sizeFactor), [int]($c.Height * $sizeFactor)) }
                    $c.Location = New-Object System.Drawing.Point([int]($c.Location.X * $sizeFactor), [int]($c.Location.Y * $sizeFactor))
                } elseif ($dock -eq [System.Windows.Forms.DockStyle]::Top -or $dock -eq [System.Windows.Forms.DockStyle]::Bottom) {
                    $c.Height = [int]($c.Height * $sizeFactor)
                } elseif ($dock -eq [System.Windows.Forms.DockStyle]::Left -or $dock -eq [System.Windows.Forms.DockStyle]::Right) {
                    $c.Width = [int]($c.Width * $sizeFactor)
                }
            }
        } catch {}
    }
}

# Modal dialogs are built at the 96-dpi baseline and shown at the current scale.
# CustomForm so a live DPI change while the dialog is open resizes its window (OS
# rect in WndProc) and fires DpiScaleChanged for the content rescale below.
function New-DialogForm($title) {
    $f = New-Object CustomForm
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.ShowInTaskbar = $false
    $f.Text = $title
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $f.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
    return $f
}
# Scale a freshly-built modal to the current DPI, theme it, dark scrollbars,
# dark title bar. Call after all controls are added and the ClientSize is set
# at the 96-dpi baseline. Children are absolute (fixed dialog, no anchors) so
# scaling the children then the client area needs no anchor arbitration.
function Complete-Dialog($dlg) {
    # Fresh leaves inherit the base font ; compensate the startup-frozen device
    # context by current/startup (= FontScale, 1.0 at startup).
    $dlgFontF = if ($script:StartupDpiFactor -gt 0) { $script:DPI_Factor / $script:StartupDpiFactor } else { 1.0 }
    Scale-ControlTree $dlg $script:DPI_Factor $dlgFontF
    $dlg.ClientSize = New-Object System.Drawing.Size([int]($dlg.ClientSize.Width * $script:DPI_Factor), [int]($dlg.ClientSize.Height * $script:DPI_Factor))
    Apply-DarkTheme $dlg
    # Plain handlers, no GetNewClosure : this script is launched via
    # [ScriptBlock]::Create(...).Invoke(), under which a closure handler cannot
    # reach a script-scope function but a plain handler can (it captures nothing).
    # The dialog + its accumulated scale therefore live in $script: singletons -
    # dialogs are modal, so only one is ever open at a time.
    $script:DlgLiveForm  = $dlg
    $script:DlgLiveScale = [float]$script:DPI_Factor
    $dlg.Add_Shown({ Set-DarkTitleBar $this; Set-DarkScrollbars $this })
    # Live DPI change while the dialog is open. Its WndProc already resized the
    # window from the OS rect ; here we rescale only the content by the ratio
    # against the dialog's own applied scale. Same frozen-DC model as the main form.
    $dlg.add_DpiScaleChanged({
        param($newScale)
        $d = $script:DlgLiveForm; if ($null -eq $d -or $d.IsDisposed) { return }
        $old = [float]$script:DlgLiveScale; if ($old -le 0) { $old = 1.0 }
        $r = [float]$newScale / $old
        if ([Math]::Abs($r - 1.0) -lt 0.001) { return }
        $script:DlgLiveScale = [float]$newScale
        $startup = [float]$script:StartupDpiFactor; if ($startup -le 0) { $startup = 1.0 }
        $d.SuspendLayout()
        try {
            [DpiContext]::FontScale = [float]$newScale / $startup
            Scale-ControlTree $d $r $r
        } finally { $d.ResumeLayout($true) }
        Set-DarkScrollbars $d
        $d.Invalidate($true)
    })
}

function Show-SimpleBookmarksTreeView {
    param($profile)
    Log "Showing bookmarks tree view for $($profile.Browser) ($($profile.Name))"
    if (-not (Test-ProfileHasBookmarks $profile)) {
        [System.Windows.Forms.MessageBox]::Show("Bookmarks not found for $($profile.Browser) / $($profile.Name)", "Not Found", 'OK', 'Error') | Out-Null
        return
    }
    $model = @(Read-ProfileBookmarks $profile)
    $dlg = New-DialogForm "Bookmarks - $($profile.Browser) ($($profile.Name))"
    $dlg.ClientSize = New-Object System.Drawing.Size(560, 460)
    $dlg.StartPosition = "CenterScreen"
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(10, 10)
    $tree.Size = New-Object System.Drawing.Size(($dlg.ClientSize.Width - 20), ($dlg.ClientSize.Height - 20))
    function Add-ViewNode($n, $parentNodes) {
        if ($n.Type -eq "Bookmark") { $tn = $parentNodes.Add($n.Text); $tn.Tag = $n.Tag }
        else { $tn = $parentNodes.Add($n.Text); foreach ($c in $n.Children) { Add-ViewNode $c $tn.Nodes } }
    }
    foreach ($n in $model) { Add-ViewNode $n $tree.Nodes }
    $dlg.Controls.Add($tree)
    Complete-Dialog $dlg
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    [void]$dlg.ShowDialog()
    $dlg.Dispose()
}

function Show-TreeView {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("restore", "backup")][string]$mode,
        $profile,                                   # backup: profile object; restore-common: $null
        [string]$source,                            # restore: folder of .url files
        [System.Windows.Forms.CheckBox]$checkbox
    )
    $bName = if ($profile) { $profile.Browser } else { "Common" }
    $pName = if ($profile) { $profile.Name } else { "Common" }
    Log "Tree view: mode=$mode, browser=$bName, profile=$pName"

    $dlg = New-DialogForm $(if ($mode -eq "restore") { "Select URL Files to Import - $bName ($pName)" } else { "Select Bookmarks - $bName ($pName)" })
    $dlg.ClientSize = New-Object System.Drawing.Size(560, 500)
    $dlg.StartPosition = "CenterParent"

    $script:suspendCheck = 0

    # Wait cursor for the tree operations (check propagation, select-all, expand).
    # UseWaitCursor keeps it across mouse moves ; Cursor.Current shows it instantly
    # even while the UI thread is busy (no message pump). Always reset in a finally
    # so it never sticks - same model as Shortcuterie.
    function Set-TreeBusy($busy) {
        if ($busy) {
            $dlg.UseWaitCursor = $true
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        } else {
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            $dlg.UseWaitCursor = $false
        }
    }

    function New-Button ($txt, $x, $y, $w, $h, $act) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $txt; $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size($w, $h); $btn.Add_Click($act)
        $dlg.Controls.Add($btn); return $btn
    }

    function Get-ChildrenCounts($node) {
        $total = 0; $checked = 0
        foreach ($child in $node.Nodes) {
            if ($child.Nodes.Count -eq 0) { $total++; if ($child.Checked) { $checked++ } }
            else { $cc = Get-ChildrenCounts $child; $total += $cc.Total; $checked += $cc.Checked }
        }
        return @{ Total = $total; Checked = $checked }
    }
    # Single post-order pass : compute (checked/total) leaf counts bottom-up and
    # set each folder's "(c/t)" suffix in ONE traversal (O(N)), writing .Text only
    # when it actually changes. Replaces the old per-folder Get-ChildrenCounts
    # recount, which re-walked each subtree for every folder (O(N^2)) and reset
    # every folder's text on every click.
    function Sync-NodeCounts($node) {
        $kids = $node.Nodes
        if ($kids.Count -eq 0) { return @{ Total = 1; Checked = ([int][bool]$node.Checked) } }
        $total = 0; $checked = 0
        foreach ($child in $kids) { $r = Sync-NodeCounts $child; $total += $r.Total; $checked += $r.Checked }
        # Strip our own " (c/t)" suffix cheaply (LastIndexOf, not a whole-text regex
        # per folder) : the suffix is always " (<digits>/<digits>)" that we appended.
        $t = $node.Text; $base = $t
        $lp = $t.LastIndexOf(' (')
        if ($lp -ge 0 -and $t[$t.Length - 1] -eq ')' -and ($t.Substring($lp + 2, $t.Length - $lp - 3) -match '^\d+/\d+$')) { $base = $t.Substring(0, $lp) }
        $newText = "$base ($checked/$total)"
        if ($t -ne $newText) { $node.Text = $newText }
        return @{ Total = $total; Checked = $checked }
    }
    function Update-AllNodeTexts($nodes) {
        foreach ($node in $nodes) { [void](Sync-NodeCounts $node) }
    }
    # Recompute ONE folder's (checked/total) from its DIRECT children's already-
    # correct displayed counts (folders) / checked state (leaves) - O(children), no
    # subtree re-walk. Used to update ancestors after a single check : the changed
    # branch is ground-truthed by Sync-NodeCounts first, siblings are already
    # correct, so summing the children is exact. Falls back to a real recount if a
    # child folder's "(c/t)" suffix is somehow missing.
    function Recount-FromChildren($node) {
        $total = 0; $checked = 0
        foreach ($child in $node.Nodes) {
            if ($child.Nodes.Count -eq 0) { $total += 1; if ($child.Checked) { $checked += 1 }; continue }
            $ct = $child.Text; $lp = $ct.LastIndexOf(' ('); $parsed = $false
            if ($lp -ge 0 -and $ct[$ct.Length - 1] -eq ')') {
                $inner = $ct.Substring($lp + 2, $ct.Length - $lp - 3); $sl = $inner.IndexOf('/')
                if ($sl -ge 0) {
                    $cc = 0; $tt = 0
                    if ([int]::TryParse($inner.Substring(0, $sl), [ref]$cc) -and [int]::TryParse($inner.Substring($sl + 1), [ref]$tt)) { $checked += $cc; $total += $tt; $parsed = $true }
                }
            }
            if (-not $parsed) { $r = Get-ChildrenCounts $child; $checked += $r.Checked; $total += $r.Total }
        }
        $t = $node.Text; $base = $t; $lp2 = $t.LastIndexOf(' (')
        if ($lp2 -ge 0 -and $t[$t.Length - 1] -eq ')' -and ($t.Substring($lp2 + 2, $t.Length - $lp2 - 3) -match '^\d+/\d+$')) { $base = $t.Substring(0, $lp2) }
        $newText = "$base ($checked/$total)"
        if ($t -ne $newText) { $node.Text = $newText }
    }
    function Update-ParentStates($nodes) {
        foreach ($n in $nodes) {
            if ($n.Nodes.Count -gt 0) { Update-ParentStates $n.Nodes; $n.Checked = $null -ne ($n.Nodes | Where-Object { $_.Checked }) }
        }
    }
    function Update-ChildNodesChecked($node, $state) {
        foreach ($child in $node.Nodes) { $child.Checked = $state; if ($child.Nodes.Count -gt 0) { Update-ChildNodesChecked $child $state } }
    }
    function Expand-CheckedParents($nodes) {
        foreach ($node in $nodes) {
            if ($node.Nodes.Count -gt 0) {
                $r = Get-ChildrenCounts $node
                if ($r.Checked -gt 0 -and $r.Checked -lt $r.Total) { $node.Expand() }
                Expand-CheckedParents $node.Nodes
            }
        }
    }
    function Update-NodesChecked($nodes, $state) {
        Set-TreeBusy $true
        $dlg.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()
        $script:suspendCheck++
        $tree.BeginUpdate()
        try {
            $queue = New-Object System.Collections.Generic.Queue[System.Windows.Forms.TreeNode]
            foreach ($n in $nodes) { $queue.Enqueue($n) }
            while ($queue.Count -gt 0) {
                $cur = $queue.Dequeue()
                $cur.Checked = $state
                foreach ($c in $cur.Nodes) { $queue.Enqueue($c) }
            }
            Update-ParentStates $tree.Nodes
            Update-AllNodeTexts $tree.Nodes
        } finally {
            $tree.EndUpdate()
            $script:suspendCheck--
            $dlg.Enabled = $true
            Set-TreeBusy $false
        }
    }

    New-Button "Select All" 10 10 100 26 { Update-NodesChecked $tree.Nodes $true } | Out-Null
    New-Button "Unselect All" 116 10 100 26 { Update-NodesChecked $tree.Nodes $false } | Out-Null
    New-Button "Expand All" ($dlg.ClientSize.Width - 206) 10 90 26 { Set-TreeBusy $true; [System.Windows.Forms.Application]::DoEvents(); try { $tree.ExpandAll() } finally { Set-TreeBusy $false } } | Out-Null
    New-Button "Collapse All" ($dlg.ClientSize.Width - 110) 10 100 26 { Set-TreeBusy $true; [System.Windows.Forms.Application]::DoEvents(); try { $tree.CollapseAll() } finally { Set-TreeBusy $false } } | Out-Null

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.CheckBoxes = $true
    $tree.Location = New-Object System.Drawing.Point(10, 46)
    $tree.Size = New-Object System.Drawing.Size(($dlg.ClientSize.Width - 20), ($dlg.ClientSize.Height - 108))

    $key = if ($mode -eq "restore" -and $null -eq $profile) { "Common|Common" } else { "$bName|$($profile.Id)" }

    if ($mode -eq "restore") {
        $isHtmlSrc = ($source -match '\.html?$' -and (Test-Path -LiteralPath $source -PathType Leaf))
        if ($isHtmlSrc) {
            # Build the tree from the HTML folder structure. Bookmark node Tag is
            # "url:<URL>" (folders keep "folder"), so no on-disk file is involved.
            $storedUrls = @()
            if ($script:RestoreSelectedUrls.ContainsKey($key)) { $storedUrls = @($script:RestoreSelectedUrls[$key] | ForEach-Object { $_.URL }) }
            $folderNodes = @{}
            foreach ($it in (ConvertFrom-BookmarkHtml ([System.IO.File]::ReadAllText($source)))) {
                $parentNode = $null; $pk = ""
                foreach ($f in $it.Folders) {
                    $pk = if ($pk -eq "") { $f } else { $pk + [char]1 + $f }
                    if (-not $folderNodes.ContainsKey($pk)) {
                        $fn = New-Object System.Windows.Forms.TreeNode($f); $fn.Tag = "folder"
                        if ($null -eq $parentNode) { [void]$tree.Nodes.Add($fn) } else { [void]$parentNode.Nodes.Add($fn) }
                        $folderNodes[$pk] = $fn
                    }
                    $parentNode = $folderNodes[$pk]
                }
                $bn = New-Object System.Windows.Forms.TreeNode($it.Name); $bn.Tag = "url:" + $it.URL
                $bn.Checked = if ($storedUrls.Count -gt 0) { ($storedUrls -contains $it.URL) } else { $true }
                if ($null -eq $parentNode) { [void]$tree.Nodes.Add($bn) } else { [void]$parentNode.Nodes.Add($bn) }
            }
        } else {
            $storedPaths = @()
            if ($script:RestoreSelectedUrls.ContainsKey($key)) { $storedPaths = $script:RestoreSelectedUrls[$key] | ForEach-Object { $_.Path } }
            function Add-RestoreNodes($path, $parent) {
                Get-ChildItem -LiteralPath $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $tn = New-Object System.Windows.Forms.TreeNode($_.Name); $tn.Tag = "folder"
                    if ($null -eq $parent) { [void]$tree.Nodes.Add($tn) } else { [void]$parent.Nodes.Add($tn) }
                    Add-RestoreNodes $_.FullName $tn
                }
                Get-ChildItem -LiteralPath $path -Filter "*.url" -File -ErrorAction SilentlyContinue | ForEach-Object {
                    $tn = New-Object System.Windows.Forms.TreeNode([System.IO.Path]::GetFileNameWithoutExtension($_.Name)); $tn.Tag = $_.FullName
                    $tn.Checked = if ($storedPaths.Count -gt 0) { ($storedPaths -contains $_.FullName) } else { $true }
                    if ($null -eq $parent) { [void]$tree.Nodes.Add($tn) } else { [void]$parent.Nodes.Add($tn) }
                }
            }
            Add-RestoreNodes $source $null
        }
    } else {
        $model = @(Read-ProfileBookmarks $profile)
        $stored = $null
        if ($script:BackupSelectedBookmarks.ContainsKey($key)) { $stored = $script:BackupSelectedBookmarks[$key] }
        function Find-Stored($storedList, $text, $isFolder, $url) {
            if (-not $storedList) { return $null }
            foreach ($s in $storedList) {
                $st = $s.Text -replace " \(\d+/\d+\)$", ""
                if ($isFolder -and $s.Type -eq "Folder" -and $st -eq $text) { return $s }
                if ((-not $isFolder) -and $s.Type -eq "Bookmark" -and $st -eq $text -and $s.Tag -eq $url) { return $s }
            }
            return $null
        }
        function Add-BackupNode($n, $parentNodes, $storedList, $forceAll) {
            if ($n.Type -eq "Bookmark") {
                $tn = $parentNodes.Add($n.Text); $tn.Tag = $n.Tag
                if ($forceAll) { $tn.Checked = $true }
                else { $s = Find-Stored $storedList $n.Text $false $n.Tag; $tn.Checked = ($null -ne $s -and $s.Checked) }
            } else {
                $tn = $parentNodes.Add($n.Text); $tn.Tag = "folder"
                if ($forceAll) { $tn.Checked = $true; foreach ($c in $n.Children) { Add-BackupNode $c $tn.Nodes $null $true } }
                else {
                    $s = Find-Stored $storedList $n.Text $true $null
                    if ($s) { $tn.Checked = $s.Checked; foreach ($c in $n.Children) { Add-BackupNode $c $tn.Nodes $s.Children $false } }
                    else { $tn.Checked = $false; foreach ($c in $n.Children) { Add-BackupNode $c $tn.Nodes $null $false } }
                }
            }
        }
        $forceAll = ($null -eq $stored)
        foreach ($n in $model) { Add-BackupNode $n $tree.Nodes $stored $forceAll }
    }

    Update-ParentStates $tree.Nodes
    Expand-CheckedParents $tree.Nodes
    Update-AllNodeTexts $tree.Nodes

    $tree.Add_AfterCheck({
        param($s, $e)
        if ($script:suspendCheck -gt 0) { return }
        $script:suspendCheck++
        # NO BeginUpdate/EndUpdate here : toggling WM_SETREDRAW while the native
        # control is still mid-processing the click leaves the just-clicked checkbox
        # unpainted until the next event (propagation appeared to land one click
        # late). Not needed anyway - the targeted update below is O(subtree + depth),
        # so its paint invalidations coalesce into a single repaint on their own.
        # Wait cursor covers the (rarer) slow case of checking a huge folder.
        Set-TreeBusy $true
        try {
            if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) { Update-ChildNodesChecked $e.Node $e.Node.Checked }
            # Ground-truth the changed subtree, then update only the ancestor chain
            # from their children's counts (O(subtree + depth), not a full-tree pass).
            [void](Sync-NodeCounts $e.Node)
            $current = $e.Node.Parent
            while ($null -ne $current) {
                $current.Checked = $null -ne ($current.Nodes | Where-Object { $_.Checked })
                Recount-FromChildren $current
                $current = $current.Parent
            }
            # Queue one repaint so every programmatically-changed checkbox refreshes
            # after the native control finishes the click (no per-node redraw storm).
            $tree.Invalidate()
        } finally { Set-TreeBusy $false; $script:suspendCheck-- }
    })
    $dlg.Controls.Add($tree)

    $btnOK = New-Button "OK" ([int]($dlg.ClientSize.Width / 2 - 60)) ($dlg.ClientSize.Height - 52) 120 40 {
        if ($mode -eq "restore") {
            function Get-CheckedFiles($nodes) {
                $result = @()
                foreach ($n in $nodes) {
                    if ($n.Checked -and $n.Tag -and $n.Tag -ne "folder") { $result += $n }
                    if ($n.Nodes.Count -gt 0) { $result += Get-CheckedFiles $n.Nodes }
                }
                return $result
            }
            $urls = @()
            foreach ($n in (Get-CheckedFiles $tree.Nodes)) {
                $tag = [string]$n.Tag
                if ($tag -like "url:*") {
                    $u = $tag.Substring(4)
                    $folders = @(); $par = $n.Parent
                    while ($null -ne $par) { $folders = @(($par.Text -replace ' \(\d+/\d+\)$', '')) + $folders; $par = $par.Parent }
                    $dir = $env:TEMP.TrimEnd('\') + "\__fbm_html_import"
                    foreach ($f in $folders) { if ($f -ne "") { $dir = Join-Path $dir (Get-ValidName $f '' $true) } }
                    $urls += @{ Name = $n.Text; URL = $u; Path = (Join-Path $dir ((Get-ValidName $n.Text $u) + ".url")) }
                } elseif (Test-Path -LiteralPath $tag) {
                    $content = Get-Content -LiteralPath $tag -Encoding UTF8 -ErrorAction SilentlyContinue
                    if ($content -and ($urlLine = $content | Where-Object { $_ -like "URL=*" })) {
                        $urls += @{ Name = $n.Text; URL = $urlLine.Substring(4).Trim(); Path = $tag }
                    }
                }
            }
            $script:RestoreSelectedUrls[$key] = $urls
            if ($null -ne $checkbox) { $checkbox.Text = "$pName ($($urls.Count) URL files)" }
            Log "Selected $($urls.Count) URL files for restore: $key"
        } else {
            function Get-CheckedNodes($nodes) {
                $results = @(); $count = 0
                foreach ($n in $nodes) {
                    $baseText = $n.Text -replace " \(\d+/\d+\)$", ""
                    if ($n.Tag -ne "folder") {
                        if ($n.Checked) { $results += @{ Text = $baseText; Tag = $n.Tag; Type = "Bookmark"; Checked = $true }; $count++ }
                    } else {
                        $cr = Get-CheckedNodes $n.Nodes
                        $results += @{ Text = $baseText; Type = "Folder"; Children = $cr.Children; Checked = $n.Checked }
                        $count += $cr.BookmarkCount
                    }
                }
                return @{ Children = $results; BookmarkCount = $count }
            }
            $result = Get-CheckedNodes $tree.Nodes
            $script:BackupSelectedBookmarks[$key] = $result.Children
            $total = Get-NodeUrlCount $model
            $checkbox.Text = "$pName ($($result.BookmarkCount) / $total)"
            $checkbox.Checked = $true
            Log "Selected $($result.BookmarkCount) bookmarks for backup: $key"
        }
        $dlg.Close()
    }
    $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    Complete-Dialog $dlg
    [void]$dlg.ShowDialog()
    $dlg.Dispose()
}

function Show-ProgressBar {
    param([string]$title, [int]$maxValue)
    Log "Progress bar: $title, max=$maxValue"
    $pf = New-DialogForm $title
    $pf.ClientSize = New-Object System.Drawing.Size(300, 70)
    $pf.StartPosition = "CenterScreen"
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(10, 20)
    $bar.Size = New-Object System.Drawing.Size(280, 22)
    $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $bar.Minimum = 0; $bar.Maximum = [Math]::Max(1, $maxValue); $bar.Value = 0
    $pf.Controls.Add($bar)
    Complete-Dialog $pf
    $pf.Show()
    return @{ Form = $pf; Bar = $bar }
}

Update-LoadingPopup 40  "Loading..."

# ============================================================================
# EXPORT
# ============================================================================
function Get-ValidName($name, $url, $isFolder = $false) {
    if (-not $isFolder -and $name -match '^(https?://)?(www\.)?') { $name = $name -replace '^(https?://)?(www\.)?', '' }
    $valid = $name -replace '[^\p{L}\p{Nd}\s@''._-]', '_' -replace '[\x00-\x1F]', '' -replace '_+', '_'
    $valid = $valid.Trim(' ._-')
    if (-not $isFolder -and ([string]::IsNullOrWhiteSpace($valid) -or $valid -match '^[\s._-]+$')) {
        $valid = $url -replace '^(https?://)?(www\.)?', '' -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
        $valid = $valid.Trim(' ._-')
        if ([string]::IsNullOrWhiteSpace($valid)) { $valid = "Unnamed_Bookmark" }
    }
    if ($valid.Length -gt 60) { $valid = $valid.Substring(0, 60).TrimEnd(' ._-') }
    return $valid
}

function Export-Bookmarks {
    Log "Starting export"
    function Get-RecursiveBookmarksCount($nodes) {
        $count = 0
        foreach ($node in $nodes) {
            if ($node.Type -eq "Bookmark" -and $node.Checked) { $count++ }
            elseif ($node.Type -eq "Folder" -and $node.Checked) { $count += Get-RecursiveBookmarksCount $node.Children }
        }
        return $count
    }
    function Convert-SelectedBookmarks($nodes, $currentPath, [ref]$currentBookmark, $progress) {
        foreach ($node in ($nodes | Where-Object { $_.Checked })) {
            if ($node.Type -eq "Bookmark") {
                $url = $node.Tag
                if ([string]::IsNullOrWhiteSpace($url) -or $url -match '^(chrome|edge|brave|opera|about|moz-extension)://') { continue }
                $validName = Get-ValidName $node.Text $url
                $filePath = Join-Path -Path $currentPath -ChildPath "$validName.url"
                if (-not (Test-Path -LiteralPath $filePath)) {
                    try {
                        [System.IO.File]::WriteAllText($filePath, "[InternetShortcut]`r`nURL=$url`r`n", (New-Object System.Text.UTF8Encoding($false)))
                    } catch {
                        Log "Failed to create URL file: $filePath. $($_.Exception.Message)"
                        $script:ExportErrors += "Failed: $filePath"
                    }
                }
                $currentBookmark.Value++
                $progress.Bar.Value = [Math]::Min($progress.Bar.Maximum, $currentBookmark.Value)
                if ($currentBookmark.Value % 20 -eq 0) { $progress.Bar.Refresh() }
            } elseif ($node.Type -eq "Folder") {
                $folderName = Get-ValidName $node.Text '' $true
                $folderPath = Join-Path -Path $currentPath -ChildPath $folderName
                try {
                    if (-not (Test-Path -LiteralPath $folderPath)) { New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null }
                    Convert-SelectedBookmarks $node.Children $folderPath $currentBookmark $progress
                } catch {
                    Log "Failed to create folder: $folderPath. $($_.Exception.Message)"
                    $script:ExportErrors += "Failed folder: $folderPath"
                }
            }
        }
    }

    $selected = @($script:BackupChecks | Where-Object { $_.Checkbox.Checked })
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile to export.", "No Profile Selected", 'OK', 'Warning') | Out-Null
        return
    }
    $exportFolder = $script:targetPath
    Log "Export target: $exportFolder"

    $totalBookmarks = 0
    foreach ($item in $selected) {
        $key = "$($item.Profile.Browser)|$($item.Profile.Id)"
        if ($script:BackupSelectedBookmarks.ContainsKey($key)) { $totalBookmarks += Get-RecursiveBookmarksCount $script:BackupSelectedBookmarks[$key] }
    }
    $totalBookmarks = [Math]::Max(1, $totalBookmarks)
    $progress = Show-ProgressBar "Exporting Bookmarks" $totalBookmarks
    $currentBookmark = 0
    try {
        foreach ($item in $selected) {
            $key = "$($item.Profile.Browser)|$($item.Profile.Id)"
            if ($script:BackupSelectedBookmarks.ContainsKey($key)) {
                Convert-SelectedBookmarks $script:BackupSelectedBookmarks[$key] $exportFolder ([ref]$currentBookmark) $progress
            }
        }
    } finally { $progress.Form.Close() }

    Reset-UI
    if ($script:ExportErrors.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Some errors occurred during export.", "Errors", 'OK', 'Error') | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("Bookmarks exported successfully.", "Done", 'OK', 'Information') | Out-Null
    }
    $script:ExportErrors = @()
}

# ============================================================================
# IMPORT
# ============================================================================
function Get-RelativePath($FullPath, $BasePath) {
    if (-not $BasePath.EndsWith("\")) { $BasePath += "\" }
    $uriFull = New-Object System.Uri($FullPath)
    $uriBase = New-Object System.Uri($BasePath)
    return [System.Uri]::UnescapeDataString($uriBase.MakeRelativeUri($uriFull).ToString()).Replace("/", "\")
}

function Get-ChromeTimestamp {
    $epoch = New-Object DateTime 1601, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
    return [long]((([DateTime]::UtcNow) - $epoch).Ticks / 10)
}

function Import-ChromiumUrls([string]$bmFile, [array]$urlObjects, [string]$baseFolder) {
    # Ensure the Bookmarks file exists.
    if (-not (Test-Path -LiteralPath $bmFile)) {
        $dir = Split-Path -LiteralPath $bmFile -Parent
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $skeleton = @{ roots = @{ bookmark_bar = @{ children = @(); type = "folder"; name = "Bookmarks bar" }; other = @{ children = @(); type = "folder" }; synced = @{ children = @(); type = "folder" } }; version = 1 }
        [System.IO.File]::WriteAllText($bmFile, ($skeleton | ConvertTo-Json -Depth 100), (New-Object System.Text.UTF8Encoding($false)))
    }
    $bmData = Get-Content -LiteralPath $bmFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $idCounter = [ref]([long](Get-ChromeTimestamp))
    function New-ChromeGuid { return ([guid]::NewGuid().ToString()) }
    function Get-ChromeFolder($parentNode, $folderName) {
        $existing = $parentNode.children | Where-Object { $_.name -eq $folderName -and $_.type -eq "folder" }
        if ($existing) { return $existing }
        $idCounter.Value++
        $newFolder = [PSCustomObject]@{
            date_added = (Get-ChromeTimestamp).ToString(); date_last_used = "0"; date_modified = (Get-ChromeTimestamp).ToString()
            guid = (New-ChromeGuid); id = $idCounter.Value.ToString(); name = $folderName; source = "user_add"; type = "folder"; children = @()
        }
        $parentNode.children += $newFolder
        return $newFolder
    }
    foreach ($obj in $urlObjects) {
        $fileDir = [System.IO.Path]::GetDirectoryName($obj.Path)
        $rel = ""
        if ($fileDir -and $baseFolder -and ($fileDir.TrimEnd('\') -ne $baseFolder.TrimEnd('\'))) { try { $rel = Get-RelativePath $fileDir $baseFolder } catch {} }
        $currentNode = $bmData.roots.bookmark_bar
        if ($rel -ne "" -and $rel -ne ".") {
            foreach ($folder in ($rel -split "\\")) { if ($folder -ne "") { $currentNode = Get-ChromeFolder $currentNode $folder } }
        }
        $exists = $currentNode.children | Where-Object { $_.type -eq "url" -and $_.url -eq $obj.URL }
        if ($exists) { continue }
        $idCounter.Value++
        $newBookmark = [PSCustomObject]@{
            date_added = (Get-ChromeTimestamp).ToString(); date_last_used = "0"; guid = (New-ChromeGuid)
            id = $idCounter.Value.ToString(); name = $obj.Name; show_icon = $false; source = "user_add"; type = "url"; url = $obj.URL
        }
        $currentNode.children += $newBookmark
    }
    if ($bmData.PSObject.Properties.Name -contains 'checksum') { $bmData.PSObject.Properties.Remove('checksum') }
    $jsonOut = $bmData | ConvertTo-Json -Depth 100 -Compress
    $tmp = "$bmFile.fbm-tmp"
    [System.IO.File]::WriteAllText($tmp, $jsonOut, (New-Object System.Text.UTF8Encoding($false)))
    if (Test-Path -LiteralPath $bmFile) { Copy-Item -LiteralPath $bmFile "$bmFile.fbm-backup" -Force }
    Move-Item -LiteralPath $tmp -Destination $bmFile -Force
}

function Test-BrowserRunning([string]$procName) {
    foreach ($p in (Get-Process -Name $procName -ErrorAction SilentlyContinue)) {
        if ($p.MainWindowHandle -ne [IntPtr]::Zero -and $p.MainWindowTitle -ne "") { return $true }
    }
    return $false
}

# Set a JSON boolean key to true with a minimal, targeted edit (no full re-
# serialization of the large Preferences file). Handles : key already present
# (flip to true), section present but key missing (insert), section absent
# (insert). Guards against a trailing comma before a closing brace.
function Set-JsonBoolTrue([string]$raw, [string]$section, [string]$key) {
    $keyPat = '"' + [regex]::Escape($key) + '"\s*:\s*(?:true|false)'
    if ([regex]::IsMatch($raw, $keyPat)) {
        return [regex]::Replace($raw, $keyPat, ('"' + $key + '":true'))
    }
    $secMatch = [regex]::Match($raw, '"' + [regex]::Escape($section) + '"\s*:\s*\{')
    if ($secMatch.Success) {
        $at = $secMatch.Index + $secMatch.Length
        $rest = $raw.Substring($at)
        $sep = if ($rest.TrimStart().StartsWith('}')) { '' } else { ',' }
        return $raw.Substring(0, $at) + '"' + $key + '":true' + $sep + $rest
    }
    $brace = $raw.IndexOf('{')
    if ($brace -lt 0) { return $raw }
    $at = $brace + 1
    $rest = $raw.Substring($at)
    $sep = if ($rest.TrimStart().StartsWith('}')) { '' } else { ',' }
    return $raw.Substring(0, $at) + '"' + $section + '":{"' + $key + '":true}' + $sep + $rest
}
# Force "always show the bookmarks bar" for a profile being imported into.
# Chromium : Preferences JSON bookmark_bar.show_on_all_tabs = true.
# Firefox  : prefs.js browser.toolbars.bookmarks.visibility = "always".
# The caller has already closed the browsers, so the files are safe to edit ;
# a one-time .fbm-backup is kept. Both files are written BOM-less UTF-8.
function Set-BookmarkBarShown($profile) {
    $dir = Split-Path -Parent $profile.Source
    if (-not $dir) { return }
    $enc = New-Object System.Text.UTF8Encoding($false)
    try {
        if ($profile.Kind -eq "Firefox") {
            $prefs = Join-Path $dir "prefs.js"
            if (-not (Test-Path -LiteralPath $prefs)) { Log "prefs.js not found : $dir"; return }
            $backup = "$prefs.fbm-backup"
            if (-not (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $prefs -Destination $backup -ErrorAction SilentlyContinue }
            $raw = [System.IO.File]::ReadAllText($prefs)
            $pat = 'user_pref\("browser\.toolbars\.bookmarks\.visibility",\s*"[^"]*"\);'
            $line = 'user_pref("browser.toolbars.bookmarks.visibility", "always");'
            if ([regex]::IsMatch($raw, $pat)) { $new = [regex]::Replace($raw, $pat, $line) }
            else { $nl = if ($raw -match "`r`n") { "`r`n" } else { "`n" }; $new = $raw.TrimEnd() + $nl + $line + $nl }
            [System.IO.File]::WriteAllText($prefs, $new, $enc)
            Log "Bookmarks toolbar forced 'always' : $($profile.Browser) / $($profile.Name)"
        } else {
            $pref = Join-Path $dir "Preferences"
            if (-not (Test-Path -LiteralPath $pref)) { Log "Preferences not found : $dir"; return }
            $backup = "$pref.fbm-backup"
            if (-not (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $pref -Destination $backup -ErrorAction SilentlyContinue }
            $raw = [System.IO.File]::ReadAllText($pref)
            $new = Set-JsonBoolTrue $raw "bookmark_bar" "show_on_all_tabs"
            if ($new -ne $raw) {
                [System.IO.File]::WriteAllText($pref, $new, $enc)
                Log "Bookmarks bar forced shown : $($profile.Browser) / $($profile.Name)"
            } else {
                Log "Bookmarks bar already shown : $($profile.Browser) / $($profile.Name)"
            }
        }
    } catch { Log "Set-BookmarkBarShown failed for $($profile.Browser)/$($profile.Name) : $($_.Exception.Message)" }
}
function Invoke-BookmarkImport($selected, $urlObjects, $baseFolder) {
    # Which browsers are involved + running (one scalar per distinct browser).
    $browsersInvolved = @($selected | ForEach-Object { $_.Browser } | Select-Object -Unique)
    $runningMap = @{}
    foreach ($bn in $browsersInvolved) {
        $def = $script:BrowserDefs | Where-Object { $_.Name -eq $bn } | Select-Object -First 1
        $runningMap[$bn] = [bool](Test-BrowserRunning $def.Proc)
    }
    $anyRunning = ($runningMap.Values | Where-Object { $_ }).Count -gt 0

    if ($anyRunning -and (-not $script:autorestore)) {
        $res = [System.Windows.Forms.MessageBox]::Show("Selected browsers will be closed and restarted. Continue?", "Confirmation", 'OKCancel')
        if ($res -ne [System.Windows.Forms.DialogResult]::OK) { Log "User canceled import"; return }
    }

    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    $progForm = Show-ProgressBar "Importing Bookmarks" 100
    function Update-Progress($v) { $progForm.Bar.Value = [Math]::Min(100, $v); $progForm.Bar.Refresh() }
    try {
        Update-Progress 10
        # Close every involved browser and wait for its process to exit.
        $procNames = @()
        foreach ($bn in $browsersInvolved) {
            $def = $script:BrowserDefs | Where-Object { $_.Name -eq $bn } | Select-Object -First 1
            Log "Closing $bn ($($def.Proc))"
            & taskkill /f /im "$($def.Proc).exe" 2>$null | Out-Null
            $procNames += $def.Proc
        }
        if ($procNames.Count -gt 0) {
            Get-Process -Name $procNames -ErrorAction SilentlyContinue | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
        }
        Update-Progress 30

        $done = 0
        foreach ($profile in $selected) {
            try {
                Log "Importing into $($profile.Browser) / $($profile.Name)"
                if ($profile.Kind -eq "Firefox") {
                    Import-FirefoxUrls $profile.Source $urlObjects $baseFolder
                } else {
                    Import-ChromiumUrls $profile.Source $urlObjects $baseFolder
                }
                if ($script:EnsureBookmarkBar) { Set-BookmarkBarShown $profile }
            } catch {
                Log "Import failed for $($profile.Browser)/$($profile.Name): $($_.Exception.Message)"
                $script:ExportErrors += "Import failed: $($profile.Browser)/$($profile.Name)"
            }
            $done++
            Update-Progress (30 + [int](50 * $done / $selected.Count))
        }
        Update-Progress 85
        Reset-UI
        Update-Progress 95

        # Restart the browsers that were running.
        foreach ($bn in $browsersInvolved) {
            if (-not $runningMap[$bn]) { continue }
            $def = $script:BrowserDefs | Where-Object { $_.Name -eq $bn } | Select-Object -First 1
            if ($def.ExePath -and (Test-Path -LiteralPath $def.ExePath)) {
                Log "Restarting $bn"
                Start-Process -FilePath $def.ExePath
            }
        }
        Update-Progress 100
    } finally {
        $progForm.Form.Close()
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    }
    Log "Import completed"
    if (-not $script:autorestore) {
        $msg = if ($script:ExportErrors.Count -gt 0) { "Import completed with some errors (see log)." } else { "Import completed successfully!" }
        [System.Windows.Forms.MessageBox]::Show($msg, "Done", 'OK', 'Information') | Out-Null
    }
    $script:ExportErrors = @()
}
function Import-bookmarks {
    Log "Starting import"
    $selected = @($script:RestoreChecks | Where-Object { $_.Checkbox.Checked } | ForEach-Object { $_.Profile })
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile to import into.", "No Profile Selected", 'OK', 'Warning') | Out-Null
        return
    }
    $base = if ($script:RestoreBaseFolder) { $script:RestoreBaseFolder } else { $script:sourcePath }
    Invoke-BookmarkImport $selected $script:CommonRestoreUrls $base
}

# ============================================================================
# HTML (Netscape bookmark file) EXPORT / IMPORT
# ============================================================================
function Write-BookmarkHtmlNodes($sb, $nodes, $indent) {
    foreach ($n in ($nodes | Where-Object { $_.Checked })) {
        if ($n.Type -eq "Bookmark") {
            $url = [string]$n.Tag
            if ([string]::IsNullOrWhiteSpace($url) -or $url -match '^(chrome|edge|brave|opera|about|moz-extension|place|javascript|data)') { continue }
            $nm = [System.Web.HttpUtility]::HtmlEncode($n.Text)
            $hu = [System.Web.HttpUtility]::HtmlEncode($url)
            [void]$sb.AppendLine("$indent<DT><A HREF=`"$hu`">$nm</A>")
        } elseif ($n.Type -eq "Folder") {
            $nm = [System.Web.HttpUtility]::HtmlEncode($n.Text)
            [void]$sb.AppendLine("$indent<DT><H3>$nm</H3>")
            [void]$sb.AppendLine("$indent<DL><p>")
            Write-BookmarkHtmlNodes $sb $n.Children "$indent    "
            [void]$sb.AppendLine("$indent</DL><p>")
        }
    }
}
function Export-BookmarksHtml([string]$outPath) {
    $selected = @($script:BackupChecks | Where-Object { $_.Checkbox.Checked })
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select at least one profile in the Backup tab first.", "Nothing selected", 'OK', 'Warning') | Out-Null
        return
    }
    $dest = [string]$outPath
    if ([string]::IsNullOrWhiteSpace($dest)) {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "Bookmarks HTML (*.html)|*.html"
        $sfd.FileName = "bookmarks.html"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $dest = $sfd.FileName
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<!DOCTYPE NETSCAPE-Bookmark-file-1>")
    [void]$sb.AppendLine('<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">')
    [void]$sb.AppendLine("<TITLE>Bookmarks</TITLE>")
    [void]$sb.AppendLine("<H1>Bookmarks</H1>")
    [void]$sb.AppendLine("<DL><p>")
    foreach ($item in $selected) {
        $key = "$($item.Profile.Browser)|$($item.Profile.Id)"
        $nodes = if ($script:BackupSelectedBookmarks.ContainsKey($key)) { $script:BackupSelectedBookmarks[$key] } else { @(Read-ProfileBookmarks $item.Profile) }
        $title = [System.Web.HttpUtility]::HtmlEncode("$($item.Profile.Browser) - $($item.Profile.Name)")
        [void]$sb.AppendLine("    <DT><H3>$title</H3>")
        [void]$sb.AppendLine("    <DL><p>")
        Write-BookmarkHtmlNodes $sb $nodes "        "
        [void]$sb.AppendLine("    </DL><p>")
    }
    [void]$sb.AppendLine("</DL><p>")
    try {
        $pdir = [System.IO.Path]::GetDirectoryName($dest)
        if ($pdir -and -not (Test-Path -LiteralPath $pdir)) { New-Item -ItemType Directory -Path $pdir -Force | Out-Null }
        [System.IO.File]::WriteAllText($dest, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        [System.Windows.Forms.MessageBox]::Show("Exported to:`n$dest", "Done", 'OK', 'Information') | Out-Null
    } catch {
        Log "HTML export failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
    }
}
# Parse a Netscape bookmark HTML into @{ Name; URL; Folders } items. The folder
# stack is tracked by <H3> (push) / </DL> (pop) : every folder's contents live in
# the <DL> right after its <H3>, and the root <DL> has no matching push, so its
# closing </DL> pops nothing (the stack is already empty by then).
function ConvertFrom-BookmarkHtml([string]$html) {
    $items = New-Object System.Collections.Generic.List[object]
    $folders = New-Object System.Collections.Generic.List[string]
    $rx = [regex]'(?is)<H3[^>]*>(?<h3>.*?)</H3>|<A[^>]+HREF="(?<url>[^"]*)"[^>]*>(?<a>.*?)</A>|</DL>'
    foreach ($m in $rx.Matches($html)) {
        if ($m.Groups['h3'].Success) {
            $folders.Add([System.Web.HttpUtility]::HtmlDecode($m.Groups['h3'].Value).Trim())
        } elseif ($m.Groups['url'].Success) {
            $url = [System.Web.HttpUtility]::HtmlDecode($m.Groups['url'].Value)
            if ([string]::IsNullOrWhiteSpace($url) -or $url -match '^(javascript|place|data):') { continue }
            $nm = [System.Web.HttpUtility]::HtmlDecode($m.Groups['a'].Value).Trim()
            [void]$items.Add(@{ Name = $nm; URL = $url; Folders = @($folders.ToArray()) })
        } else {
            if ($folders.Count -gt 0) { $folders.RemoveAt($folders.Count - 1) }
        }
    }
    return $items
}
# Read the restore source (a folder of .url files, or a .html bookmarks file)
# into a common URL set. HTML items get a synthetic Path under a virtual base so
# Import-ChromiumUrls / Import-FirefoxUrls rebuild the folder structure exactly as
# they do for a .url-file source. Returns @{ Urls = @(...); Base = <folder> }.
function Get-SourceUrls {
    $src = [string]$script:sourcePath
    if ($src -match '\.html?$' -and (Test-Path -LiteralPath $src -PathType Leaf)) {
        $base = $env:TEMP.TrimEnd('\') + "\__fbm_html_import"
        $urls = @()
        foreach ($it in (ConvertFrom-BookmarkHtml ([System.IO.File]::ReadAllText($src)))) {
            $dir = $base
            foreach ($f in $it.Folders) { if ($f -ne "") { $dir = Join-Path $dir (Get-ValidName $f '' $true) } }
            $urls += @{ Name = $it.Name; URL = $it.URL; Path = (Join-Path $dir ((Get-ValidName $it.Name $it.URL) + ".url")) }
        }
        return @{ Urls = $urls; Base = $base }
    }
    $urls = @()
    if ($src -and (Test-Path -LiteralPath $src -PathType Container)) {
        foreach ($file in (Get-ChildItem -LiteralPath $src -Filter "*.url" -Recurse -ErrorAction SilentlyContinue)) {
            $content = Get-Content -LiteralPath $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
            if ($content -and ($urlLine = $content | Where-Object { $_ -like "URL=*" })) {
                $urls += @{ Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name); URL = $urlLine.Substring(4).Trim(); Path = $file.FullName }
            }
        }
    }
    return @{ Urls = $urls; Base = $src }
}

# ============================================================================
# DATA-DRIVEN UI
# ============================================================================
# Build a FlowLayoutPanel of one GroupBox per detected browser, each with per-profile rows.
# Returns @{ Panel; Checks } where Checks = list of @{ Profile; Checkbox }.
function New-BrowserFlow([string]$mode) {
    $checks = New-Object System.Collections.ArrayList
    $btnText = if ($mode -eq "backup") { "Select" } else { "View" }
    $built = New-Object System.Collections.ArrayList   # @{ Group; Height } per browser

    foreach ($def in $script:DetectedBrowsers) {
        $profiles = @(Get-BrowserProfileList $def)
        $group = New-Object System.Windows.Forms.GroupBox
        $group.Text = $def.Name
        $group.AutoSize = $true
        $group.AutoSizeMode = "GrowAndShrink"
        $group.MinimumSize = New-Object System.Drawing.Size(258, 60)
        $group.Margin = New-Object System.Windows.Forms.Padding(5)
        $group.Padding = New-Object System.Windows.Forms.Padding(6, 4, 6, 6)
        if ($script:IsDark) { $group.Add_Paint($script:GroupPaintHandler) }
        $y = 24
        if ($profiles.Count -eq 0) {
            $l = New-Object System.Windows.Forms.Label
            $l.Text = "No profiles found"; $l.AutoSize = $true
            $l.Location = New-Object System.Drawing.Point(12, $y)
            $group.Controls.Add($l)
        }
        foreach ($p in $profiles) {
            $has = Test-ProfileHasBookmarks $p
            $count = if ($has -and $mode -eq "backup") { Get-NodeUrlCount (Read-ProfileBookmarks $p) } else { 0 }
            # A profile with no Bookmarks file can still be a RESTORE target for
            # Chromium / Opera GX : Import-ChromiumUrls creates the file. Firefox
            # needs an existing places.sqlite, so it can't be seeded from scratch.
            $canRestoreInto = ($p.Kind -ne "Firefox")
            $sfx = if ($mode -eq "backup") {
                if ($has -and $count -gt 0) { " ($count / $count)" } else { " (empty)" }
            } else {
                if ($has) { "" } elseif ($canRestoreInto) { " (empty)" } else { " (not available)" }
            }
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $p.Name + $sfx
            $chk.AutoSize = $false
            $chk.Size = New-Object System.Drawing.Size(190, 22)
            $chk.Location = New-Object System.Drawing.Point(12, $y)
            $chk.Enabled = if ($mode -eq "backup") { ($has -and $count -gt 0) } else { ($has -or $canRestoreInto) }
            $btn = New-Object System.Windows.Forms.Button
            $btn.Text = $btnText
            $btn.Size = New-Object System.Drawing.Size(52, 22)
            $btn.Location = New-Object System.Drawing.Point(206, $y)
            $btn.Enabled = $has -and ($mode -eq "restore" -or $count -gt 0)
            $btn.Tag = @{ Profile = $p; Checkbox = $chk; Mode = $mode }
            $btn.Add_Click({
                $t = $this.Tag
                [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
                $this.Enabled = $false
                try {
                    if ($t.Mode -eq "backup") { Show-TreeView -mode "backup" -profile $t.Profile -checkbox $t.Checkbox }
                    else { Show-SimpleBookmarksTreeView $t.Profile }
                } finally { $this.Enabled = $true; [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }
            })
            if ($mode -eq "backup" -and $has -and $count -gt 0) {
                $script:BackupSelectedBookmarks["$($p.Browser)|$($p.Id)"] = Read-ProfileBookmarks $p
            }
            $group.Controls.Add($chk)
            $group.Controls.Add($btn)
            [void]$checks.Add(@{ Profile = $p; Checkbox = $chk })
            $y += 28
        }
        [void]$built.Add(@{ Group = $group; Height = (34 + 28 * [Math]::Max(1, $profiles.Count)) })
    }
    # Two INDEPENDENT vertical columns (masonry : each column packs top-to-bottom
    # on its own, no cross-column row alignment), balanced greedily by estimated
    # height, hosted in a ScrollPanel (fixed always-present vertical scrollbar, so
    # the content width is constant -> exactly two columns, no scrollbar tipping).
    $colL = New-Object System.Windows.Forms.FlowLayoutPanel
    $colL.FlowDirection = "TopDown"; $colL.WrapContents = $false; $colL.AutoSize = $true
    $colL.AutoSizeMode = "GrowAndShrink"; $colL.Margin = New-Object System.Windows.Forms.Padding(0)
    $colR = New-Object System.Windows.Forms.FlowLayoutPanel
    $colR.FlowDirection = "TopDown"; $colR.WrapContents = $false; $colR.AutoSize = $true
    $colR.AutoSizeMode = "GrowAndShrink"; $colR.Margin = New-Object System.Windows.Forms.Padding(0)
    $hL = 0; $hR = 0
    foreach ($b in $built) {
        if ($hL -le $hR) { $colL.Controls.Add($b.Group); $hL += $b.Height }
        else { $colR.Controls.Add($b.Group); $hR += $b.Height }
    }
    $rowFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $rowFlow.FlowDirection = "LeftToRight"; $rowFlow.WrapContents = $false; $rowFlow.AutoSize = $true
    $rowFlow.AutoSizeMode = "GrowAndShrink"; $rowFlow.Margin = New-Object System.Windows.Forms.Padding(0)
    $rowFlow.Location = New-Object System.Drawing.Point(0, 0)
    $rowFlow.Controls.Add($colL); $rowFlow.Controls.Add($colR)
    $sp = New-Object ScrollPanel
    $sp.Dock = "Fill"
    if ($script:IsDark) { $sp.BackColor = $script:Theme.Back }
    $sp.Controls.Add($rowFlow)
    return @{ Panel = $sp; Checks = $checks }
}

function Reset-UI {
    Log "Resetting UI"
    $script:BackupSelectedBookmarks = @{}
    $script:RestoreSelectedUrls = @{}
    $script:CommonRestoreUrls = @{}
    $script:ExportErrors = @()

    # Freshly built leaves inherit the base container font ; scale them to the
    # current physical size on the startup-frozen device context (= FontScale,
    # 1.0 at startup).
    $rebuildFontF = if ($script:StartupDpiFactor -gt 0) { $script:DPI_Factor / $script:StartupDpiFactor } else { 1.0 }

    $backupTabHost.Controls.Clear()
    $backupBuild = New-BrowserFlow "backup"
    $script:BackupChecks = $backupBuild.Checks
    Scale-ControlTree $backupBuild.Panel $script:DPI_Factor $rebuildFontF
    $backupTabHost.Controls.Add($backupBuild.Panel)

    $restoreTabHost.Controls.Clear()
    $restoreBuild = New-BrowserFlow "restore"
    $script:RestoreChecks = $restoreBuild.Checks
    Scale-ControlTree $restoreBuild.Panel $script:DPI_Factor $rebuildFontF
    $restoreTabHost.Controls.Add($restoreBuild.Panel)

    # Common URL set from the source : a folder of .url files, or a .html file.
    $src = Get-SourceUrls
    $script:CommonRestoreUrls = $src.Urls
    $script:RestoreBaseFolder = $src.Base
    $script:RestoreSelectedUrls["Common|Common"] = $src.Urls
    $commonSelectButton.Text = "Select URLs to import ($(@($src.Urls).Count))"

    Apply-DarkTheme $form
    Set-DarkScrollbars $backupTabHost
    Set-DarkScrollbars $restoreTabHost
}

Update-LoadingPopup 50  "Loading..."

##############################################
# MAIN INTERFACE
##############################################

$form = New-Object CustomForm
$form.SuspendLayout()
$form.Text = "Fast Bookmarks Manager"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.ClientSize = New-Object System.Drawing.Size(582, 510)   # 96-dpi baseline
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
if ($script:IsDark) { $form.BackColor = $script:Theme.Back }

# TabControl fills the client area ; the Refresh button floats over the empty
# right end of the tab strip (it does not occupy its own row).
$tabControl = New-Object DarkTabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Size = New-Object System.Drawing.Size(72, 22)
$refreshButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 78), 4)
$refreshButton.Add_Click({
    $this.Enabled = $false
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try { Reset-UI } finally { $this.Enabled = $true; [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }
})
$form.Controls.Add($refreshButton)

Update-LoadingPopup 60  "Loading..."

# Center a horizontal row of buttons in a Dock=Bottom/Top panel (scaled gap).
function Center-ButtonRow($panel, $buttons) {
    $gap = [int](10 * $script:DPI_Factor)
    $total = 0
    foreach ($b in $buttons) { $total += $b.Width }
    $total += $gap * [Math]::Max(0, $buttons.Count - 1)
    $x = [int](($panel.ClientSize.Width - $total) / 2)
    foreach ($b in $buttons) {
        $b.Location = New-Object System.Drawing.Point($x, [int](($panel.ClientSize.Height - $b.Height) / 2))
        $x += $b.Width + $gap
    }
}
# Lay out a "folder / file" row (label + textbox + Open + Browse) from the panel
# width, plus an optional centered extra button below it.
function Layout-FolderRow($panel, $label, $textbox, $openBtn, $browseBtn, $extraBtn, $extraChk) {
    $cw = $panel.ClientSize.Width
    $sc = [float]$script:DPI_Factor; if ($sc -le 0) { $sc = 1.0 }
    $m = [int](12 * $sc); $gap = [int](6 * $sc)
    $label.Location = New-Object System.Drawing.Point($m, [int](8 * $sc))
    $y2 = [int](32 * $sc)
    $browseBtn.Location = New-Object System.Drawing.Point(($cw - $m - $browseBtn.Width), $y2)
    $openBtn.Location = New-Object System.Drawing.Point(($browseBtn.Location.X - $gap - $openBtn.Width), $y2)
    $textbox.Location = New-Object System.Drawing.Point($m, $y2)
    $textbox.Width = [Math]::Max(60, $openBtn.Location.X - $gap - $m)
    $ey = [int](64 * $sc)
    if ($null -ne $extraBtn -and $null -ne $extraChk) {
        # Centre the button + checkbox as one pair on the same line.
        $pairGap = [int](12 * $sc)
        $total = $extraBtn.Width + $pairGap + $extraChk.Width
        $x = [int](($cw - $total) / 2); if ($x -lt $m) { $x = $m }
        $extraBtn.Location = New-Object System.Drawing.Point($x, $ey)
        $extraChk.Location = New-Object System.Drawing.Point(($x + $extraBtn.Width + $pairGap), ($ey + [int](($extraBtn.Height - $extraChk.Height) / 2)))
    } elseif ($null -ne $extraBtn) {
        $extraBtn.Location = New-Object System.Drawing.Point([int](($cw - $extraBtn.Width) / 2), $ey)
    }
}
# Open a folder (create if missing), or reveal a file in Explorer.
function Open-PathLocation($p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return }
    if (Test-Path -LiteralPath $p -PathType Leaf) { Start-Process explorer.exe "/select,`"$p`""; return }
    if (-not (Test-Path -LiteralPath $p)) { try { New-Item -ItemType Directory -Path $p -Force | Out-Null } catch {} }
    if (Test-Path -LiteralPath $p) { Start-Process -FilePath $p }
    else { [System.Windows.Forms.MessageBox]::Show("Not found:`n$p", "Not Found", 'OK', 'Warning') | Out-Null }
}

# --- Backup tab --- (folder/HTML target row on top, single Export button below)
$tabBackup = New-Object System.Windows.Forms.TabPage
$tabBackup.Text = "Backup"
$backupTop = New-Object System.Windows.Forms.Panel
$backupTop.Dock = "Top"; $backupTop.Height = 64
$backupSaveLabel = New-Object System.Windows.Forms.Label
$backupSaveLabel.Text = "Where to save backups (a folder, or a .html file) :"
$backupSaveLabel.AutoSize = $true
$backupTop.Controls.Add($backupSaveLabel)
$targetFolderTextBox = New-Object System.Windows.Forms.TextBox
$targetFolderTextBox.Size = New-Object System.Drawing.Size(300, 24)
$targetFolderTextBox.Text = $script:targetPath
$backupTop.Controls.Add($targetFolderTextBox)
$openTargetButton = New-Object System.Windows.Forms.Button
$openTargetButton.Text = "Open"; $openTargetButton.Size = New-Object System.Drawing.Size(60, 24)
$openTargetButton.Add_Click({ Open-PathLocation $script:targetPath })
$backupTop.Controls.Add($openTargetButton)
$browseTargetButton = New-Object System.Windows.Forms.Button
$browseTargetButton.Text = "Browse..."; $browseTargetButton.Size = New-Object System.Drawing.Size(80, 24)
$browseTargetButton.Add_Click({
    $sp = ([FolderSelectDialog]::new($targetFolderTextBox.Text, "Select Backup Target Folder", "Select the folder to store backups.")).getPath()
    if ($sp -ne "") { $targetFolderTextBox.Text = $sp }
})
$backupTop.Controls.Add($browseTargetButton)
$backupTop.Add_Resize({ Layout-FolderRow $backupTop $backupSaveLabel $targetFolderTextBox $openTargetButton $browseTargetButton $null })

$backupBottom = New-Object System.Windows.Forms.Panel
$backupBottom.Dock = "Bottom"; $backupBottom.Height = 46
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export"
$exportButton.Size = New-Object System.Drawing.Size(160, 34)
$exportButton.Anchor = "None"
$exportButton.Add_Click({
    $this.Enabled = $false
    try {
        $t = [string]$script:targetPath
        if ($t -match '\.html?$') { Export-BookmarksHtml $t } else { Export-Bookmarks }
    } finally { $this.Enabled = $true }
})
$backupBottom.Add_Resize({ Center-ButtonRow $backupBottom @($exportButton) })
$backupBottom.Controls.Add($exportButton)

$backupTabHost = New-Object System.Windows.Forms.Panel
$backupTabHost.Dock = "Fill"
$tabBackup.Controls.Add($backupTabHost)
$tabBackup.Controls.Add($backupBottom)
$tabBackup.Controls.Add($backupTop)

Update-LoadingPopup 70  "Loading..."

# --- Restore tab --- (folder/HTML source row + Select URLs on top, single Import)
$tabRestore = New-Object System.Windows.Forms.TabPage
$tabRestore.Text = "Restore"
$restoreTop = New-Object System.Windows.Forms.Panel
$restoreTop.Dock = "Top"; $restoreTop.Height = 100
$restoreSourceLabel = New-Object System.Windows.Forms.Label
$restoreSourceLabel.Text = "Where to restore backups from (Folder / HTML) :"
$restoreSourceLabel.AutoSize = $true
$restoreTop.Controls.Add($restoreSourceLabel)
$sourceFolderTextBox = New-Object System.Windows.Forms.TextBox
$sourceFolderTextBox.Size = New-Object System.Drawing.Size(300, 24)
$sourceFolderTextBox.Text = $script:sourcePath
$restoreTop.Controls.Add($sourceFolderTextBox)
$openSourceButton = New-Object System.Windows.Forms.Button
$openSourceButton.Text = "Open"; $openSourceButton.Size = New-Object System.Drawing.Size(60, 24)
$openSourceButton.Add_Click({ Open-PathLocation $script:sourcePath })
$restoreTop.Controls.Add($openSourceButton)
$browseSourceButton = New-Object System.Windows.Forms.Button
$browseSourceButton.Text = "Browse..."; $browseSourceButton.Size = New-Object System.Drawing.Size(80, 24)
$browseSourceButton.Add_Click({
    $sp = ([FolderSelectDialog]::new($sourceFolderTextBox.Text, "Select Restore Source Folder", "Select the folder to restore backups from.")).getPath()
    if ($sp -ne "") { $sourceFolderTextBox.Text = $sp; Reset-UI }
})
$restoreTop.Controls.Add($browseSourceButton)
$commonSelectButton = New-Object System.Windows.Forms.Button
$commonSelectButton.Size = New-Object System.Drawing.Size(190, 30)
$commonSelectButton.Text = "Select URLs to import (0)"
$commonSelectButton.Anchor = "None"
$commonSelectButton.Add_Click({
    $this.Enabled = $false
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Show-TreeView -mode "restore" -profile $null -source $script:sourcePath -checkbox $null
        if ($script:RestoreSelectedUrls.ContainsKey("Common|Common")) {
            $cnt = ($script:RestoreSelectedUrls["Common|Common"]).Count
            $commonSelectButton.Text = "Select URLs to import ($cnt)"
            $script:CommonRestoreUrls = $script:RestoreSelectedUrls["Common|Common"]
        }
    } finally { $this.Enabled = $true; [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }
})
$restoreTop.Controls.Add($commonSelectButton)
$ensureBarCheckbox = New-Object System.Windows.Forms.CheckBox
$ensureBarCheckbox.Text = "Ensure bookmarks bar is always shown"
$ensureBarCheckbox.AutoSize = $true
$ensureBarCheckbox.Checked = $false
$ensureBarCheckbox.Anchor = "None"
$ensureBarCheckbox.Add_CheckedChanged({ $script:EnsureBookmarkBar = $ensureBarCheckbox.Checked })
$restoreTop.Controls.Add($ensureBarCheckbox)
$restoreTop.Add_Resize({ Layout-FolderRow $restoreTop $restoreSourceLabel $sourceFolderTextBox $openSourceButton $browseSourceButton $commonSelectButton $ensureBarCheckbox })

$restoreBottom = New-Object System.Windows.Forms.Panel
$restoreBottom.Dock = "Bottom"; $restoreBottom.Height = 46
$importButton = New-Object System.Windows.Forms.Button
$importButton.Text = "Import"
$importButton.Size = New-Object System.Drawing.Size(160, 34)
$importButton.Anchor = "None"
$importButton.Add_Click({
    $this.Enabled = $false
    try { Import-bookmarks } finally { $this.Enabled = $true }
})
$restoreBottom.Add_Resize({ Center-ButtonRow $restoreBottom @($importButton) })
$restoreBottom.Controls.Add($importButton)

$restoreTabHost = New-Object System.Windows.Forms.Panel
$restoreTabHost.Dock = "Fill"
$tabRestore.Controls.Add($restoreTabHost)
$tabRestore.Controls.Add($restoreBottom)
$tabRestore.Controls.Add($restoreTop)

$targetFolderTextBox.Add_TextChanged({ $script:targetPath = $targetFolderTextBox.Text })
# Editable source : update the path per keystroke (cheap) but only rebuild the UI
# when editing finishes (Leave) or a folder is picked (Browse) - Reset-UI re-reads
# every profile + the source, too heavy to run on each keystroke.
$sourceFolderTextBox.Add_TextChanged({ $script:sourcePath = $sourceFolderTextBox.Text })
$sourceFolderTextBox.Add_Leave({ Reset-UI })

$tabControl.TabPages.Add($tabBackup)
$tabControl.TabPages.Add($tabRestore)
$form.Controls.Add($tabControl)
$refreshButton.BringToFront()
# A flow panel on a not-yet-shown tab has no handle at startup, so its scrollbar
# stays default-light. Re-theme the scrollbars of whichever tab becomes visible.
$tabControl.Add_SelectedIndexChanged({
    try { if ($null -ne $tabControl.SelectedTab) { Set-DarkScrollbars $tabControl.SelectedTab } } catch {}
})

Update-LoadingPopup 80  "Loading..."

# Startup scaling : the static controls are built at the 96-dpi baseline. Scale
# their bounds (fonts stay : a control created now renders at the monitor DPI),
# grow the fixed window, and re-place the floating Refresh button. The dynamic
# browser flow panels are built + scaled inside Reset-UI, so the static tree is
# scaled FIRST while their hosts are empty (no double-scaling).
[DpiContext]::Scale = [float]$script:DPI_Factor
[DpiContext]::FontScale = 1.0
try { $tabControl.ApplyDpiScaling() } catch {}
Scale-ControlTree $form $script:DPI_Factor 1.0
$form.ClientSize = New-Object System.Drawing.Size([int](582 * $script:DPI_Factor), [int](510 * $script:DPI_Factor))
$refreshButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $refreshButton.Width - [int](6 * $script:DPI_Factor)), [int](4 * $script:DPI_Factor))
$global:FbmAppliedScale = [float]$script:DPI_Factor

$script:DetectedBrowsers = @(Get-DetectedBrowsers)
Reset-UI
Apply-DarkTheme $form

# Live per-monitor DPI change. Under the PowerShell host the control device
# contexts stay FROZEN at the startup DPI, so both bounds AND fonts scale by the
# new/old ratio to keep their physical size. Only LEAF fonts are walked (see
# Scale-ControlTree) ; container-drawn text (tab labels, group titles) self-
# scales through DpiContext.FontScale = current/startup, updated here. The
# applied scale is tracked in $global: (reliable across event-handler
# invocations) + a re-entrancy flag so duplicate/nested WM_DPICHANGED deliveries
# can't compound.
$form.add_DpiScaleChanged({
    param($newScale)
    if ($global:FbmDpiBusy) { return }
    $oldScale = $global:FbmAppliedScale
    if ($null -eq $oldScale -or $oldScale -le 0) { $oldScale = [float]$script:StartupDpiFactor }
    if ($null -eq $oldScale -or $oldScale -le 0) { $oldScale = 1.0 }
    $ratio = [float]$newScale / [float]$oldScale
    if ($ratio -le 0) { $ratio = 1 }
    if ([Math]::Abs($ratio - 1.0) -lt 0.001) { return }
    $startup = [float]$script:StartupDpiFactor; if ($startup -le 0) { $startup = 1.0 }
    $global:FbmDpiBusy = $true
    $form.SuspendLayout()
    try {
        $script:DPI_Factor = [float]$newScale
        $global:FbmAppliedScale = [float]$newScale
        [DpiContext]::Scale = [float]$newScale
        [DpiContext]::FontScale = [float]$newScale / $startup
        try { $tabControl.ApplyDpiScaling() } catch {}
        # The window (and its Dock=Fill tab control) was already resized by the
        # WndProc from the OS-suggested rect ; here we only rescale the controls.
        Scale-ControlTree $form $ratio $ratio
        $refreshButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $refreshButton.Width - [int](6 * $newScale)), [int](4 * $newScale))
        # Re-center the floating action button rows in their re-sized host panels
        # (deterministic : after the walk has scaled both buttons and panel).
        try { Center-ButtonRow $backupBottom @($exportButton) } catch {}
        try { Center-ButtonRow $restoreBottom @($importButton) } catch {}
        try { Layout-FolderRow $backupTop $backupSaveLabel $targetFolderTextBox $openTargetButton $browseTargetButton $null } catch {}
        try { Layout-FolderRow $restoreTop $restoreSourceLabel $sourceFolderTextBox $openSourceButton $browseSourceButton $commonSelectButton $ensureBarCheckbox } catch {}
    } finally {
        $form.ResumeLayout($true)
        $global:FbmDpiBusy = $false
    }
    Set-DarkScrollbars $form
    $form.Invalidate($true)
})

$form.Add_Load({
    Update-LoadingPopup 90 "Finalizing..."
})

$form.Add_Shown({ Set-DarkTitleBar $form; Set-DarkScrollbars $form })

if ($script:autorestore) {
    Log "Auto-restore mode enabled"
    $form.Add_Shown({
        $form.BeginInvoke([Action]{
            Log "Starting auto-restore sequence"
            $tabControl.SelectedTab = $tabRestore
            $defaultsFound = $false
            foreach ($item in $script:RestoreChecks) {
                if ($item.Checkbox.Enabled -and $item.Profile.Id -match '^(Default|default)') {
                    $item.Checkbox.Checked = $true; $defaultsFound = $true
                }
            }
            if (-not $defaultsFound) { Log "No default profiles found"; $form.Close(); return }
            Import-bookmarks
            Log "Auto-restore completed"
            $form.Close()
        })
    })
}

$form.ResumeLayout($true)
Close-LoadingPopup
[System.Windows.Forms.Application]::Run($form)
try { $form.Dispose() } catch {}
Log "--- Application closed ---"
