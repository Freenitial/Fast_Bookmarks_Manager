<# ::
    cls & @echo off & chcp 437 >nul & title Fast Bookmarks Manager

    if /i "%~1"=="/?"       goto :help
    if /i "%~1"=="-?"       goto :help
    if /i "%~1"=="--?"      goto :help
    if /i "%~1"=="/help"    goto :help
    if /i "%~1"=="-help"    goto :help
    if /i "%~1"=="--help"   goto :help

    copy /y "%~f0" "%TEMP%\%~n0.ps1" >NUL && powershell -Nologo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%TEMP%\%~n0.ps1" %*
    for %%A in (%*) do if /I "%%A"=="-autorestore" exit
    exit /b

    :help
    mode con: cols=128 lines=60
    echo.
    echo.
    echo    =============================================================================
    echo                              Fast Bookmarks Manager v0.9
    echo                                          ---
    echo                     Author : Leo Gillet - Freenitial on GitHub
    echo    =============================================================================
    echo.
    echo.
    echo    DESCRIPTION:
    echo       -----------
    echo       Fast Bookmarks Manager is a PowerShell tool with a graphical interface
    echo       for exporting and importing bookmarks-bar from Chrome and Edge profiles.
    echo.
    echo       Key features:
    echo         - Export bookmarks-bar as .url files and folder structure organized by browser and profile
    echo         - Import bookmarks from exported .url files
    echo         - Manage target folder for backup, and source folder for restoration
    echo         - Intuitive tree view for bookmarks selection
    echo         - Auto-restore for each default profile of each browser with the -autorestore parameter
    echo         - With Auto-restore you can add argument -source "C:\Folder_containing_url_files\" if not in script dir
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
    echo          start "" /d "SCRIPT_DIR" Fast_Bookmarks_Manager ^
    echo                                   -autorestore ^
    echo                                   -source "C:\Path\To\Source" ^
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
}
"@ -Language CSharp -ReferencedAssemblies @("System.Runtime.InteropServices", "System.Windows.Forms", "System.ComponentModel.Primitives")
[DPIHelper]::SetProcessDpiAwarenessContext([DPIHelper]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2) | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ignoreAfterCheck = $false
$script:BackupSelectedBookmarks = @{}      # Dictionary: "Browser|Profile" => stored tree selection (with Checked states)
$script:RestoreSelectedUrls = @{}          # Dictionary: "Browser|Profile" => URL objects for restore
$script:CommonRestoreUrls = @{}            # Dictionary key "Common|Common" for common URL selection
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
    while ((Test-Path -LiteralPath $newPath) -and (-not (Test-FileWritable -Path $newPath))) {
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
if (-not (Test-FileWritable -Path $logfile)) {
    try { Add-Content -Path $logfile -Value $message -ErrorAction Stop } 
    catch {
        Write-Host "File is locked or read-only. Searching for an available logfile."
        $logfile = Get-AvailableLogFileName -FilePath $logfile
        Write-Host "New logfile: $logfile"
    }
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

if ($source.IsPresent) {
    if (Test-PathAccess -Path $source) { $script:sourcePath = $source }
    else {
        Log "Source path access failed. Using current location."
        $script:sourcePath = (Get-Location).Path
    }
}

function Get-BrowserProfiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("chrome", "edge")]
        [string]$Browser
    )
    if ($Browser -eq "chrome") {
        Log "Getting Chrome profiles"
        $browserData = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    }
    elseif ($Browser -eq "edge") {
        Log "Getting Edge profiles"
        $browserData = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
    }
    if (-not (Test-Path $browserData)) { 
        Log "Data folder not found: $browserData"
        return $null 
    }
    $BrowserProfiles = @("Default")
    $BrowserProfiles += Get-ChildItem -Path $browserData -Directory | Where-Object { $_.Name -match "^Profile \d+$" } | ForEach-Object { $_.Name }
    Log "Found $($BrowserProfiles.Count) profiles for $Browser"
    return $BrowserProfiles
}

function Get-BookmarksCount ($bookmarksFile) {
    if (-not (Test-Path $bookmarksFile)) { return 0 }
    $json = Get-Content -Path $bookmarksFile -Raw -Encoding UTF8 | ConvertFrom-Json
    function Get-RecursiveCount ($node) {
        $count = 0
        if ($node.type -eq 'url') { $count++ }
        if ($node.children) { foreach ($child in $node.children) { $count += Get-RecursiveCount $child } }
        return $count
    }
    return Get-RecursiveCount $json.roots.bookmark_bar
}

function Show-SimpleBookmarksTreeView {
    param([string]$browser, [string]$BrowserProfile, [string]$bookmarksFile)
    Log "Showing simple bookmarks tree view for $browser ($BrowserProfile)"
    if (-not (Test-Path $bookmarksFile)) {
        Log "Bookmarks file not found: $bookmarksFile"
        [System.Windows.Forms.MessageBox]::Show("Bookmarks file not found: $bookmarksFile", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $json = Get-Content -Path $bookmarksFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $ReadBookmarksForm = New-Object System.Windows.Forms.Form
    $ReadBookmarksForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $ReadBookmarksForm.MaximizeBox = $false
    $ReadBookmarksForm.MinimizeBox = $false
    $ReadBookmarksForm.ShowInTaskbar = $false
    $ReadBookmarksForm.Text = "Bookmarks - $browser ($BrowserProfile)"
    $ReadBookmarksForm.Size = New-Object System.Drawing.Size($form.Width, $($form.Height - 50))
    $ReadBookmarksForm.StartPosition = "CenterScreen"
    $ReadBookmarksForm.Font = New-Object System.Drawing.Font("Arial",10)
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Location = New-Object System.Drawing.Point(10,10)
    $tree.Size = New-Object System.Drawing.Size($($ReadBookmarksForm.Width - 35), $($ReadBookmarksForm.Height - 60))
    $tree.Anchor = "Top,Left,Bottom,Right"

    function Add-Node ($node, $parent) {
        if ($node.type -eq 'url') {
            $name = if ([string]::IsNullOrWhiteSpace($node.name)) { $node.url } else { [System.Web.HttpUtility]::HtmlDecode($node.name) }
            $tn = $parent.Nodes.Add($name)
            $tn.Tag = $node.url
        }
        if ($node.children) {
            $name = if ([string]::IsNullOrWhiteSpace($node.name)) { "Folder" } else { [System.Web.HttpUtility]::HtmlDecode($node.name) }
            $tn = $parent.Nodes.Add($name)
            foreach ($child in $node.children) { Add-Node $child $tn }
        }
    }
    foreach ($child in $json.roots.bookmark_bar.children) { Add-Node $child $tree }
    $ReadBookmarksForm.Controls.Add($tree)
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor
    $ReadBookmarksForm.ShowDialog()
}

function Show-TreeView {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("restore", "backup")]
        [string]$mode,
        [string]$browser,
        [string]$BrowserProfile,
        [string]$source,
        [System.Windows.Forms.CheckBox]$checkbox
    )
    Log "Showing tree view: mode=$mode, browser=$browser, profile=$BrowserProfile, source=$source"
    
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Text = if ($mode -eq "restore") { "Select URL Files to Import - $browser ($BrowserProfile)" } else { "Select Bookmarks - $browser ($BrowserProfile)" }
    $dlg.Size = New-Object System.Drawing.Size($form.Width, $($form.Height - 50))
    $dlg.StartPosition = "CenterParent"
    $dlg.Font = New-Object System.Drawing.Font("Arial", 10)

    # Helper function for create buttons
    function New-Button ($txt, $x, $y, $w, $h, $act) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $txt; $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size($w, $h); $btn.Add_Click($act)
        $dlg.Controls.Add($btn); return $btn
    }
    
    # Recursive check/uncheck function
    function Update-NodesChecked ($nodes, $state) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $script:ignoreAfterCheck = $true
        try {
            $nodesToProcess = New-Object System.Collections.Generic.Queue[System.Windows.Forms.TreeNode]
            foreach ($n in $nodes) { $nodesToProcess.Enqueue($n) }
            while ($nodesToProcess.Count -gt 0) {
                $currentNode = $nodesToProcess.Dequeue()
                $currentNode.Checked = $state
                foreach ($childNode in $currentNode.Nodes) { $nodesToProcess.Enqueue($childNode) }
                if ($nodesToProcess.Count % 100 -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
            }
            Update-ParentStates $tree.Nodes
            Update-AllNodeTexts $tree.Nodes
        }
        finally {
            $script:ignoreAfterCheck = $false
            $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor 
        }
    }
    
    function Get-ChildrenCounts {
        param($node)
        $totalChildren = 0
        $checkedChildren = 0
        foreach ($child in $node.Nodes) {
            if ($child.Nodes.Count -eq 0) {
                $totalChildren++
                if ($child.Checked) { $checkedChildren++ }
            } else {
                # For folders, don't count the node itself, only its children
                $childCounts = Get-ChildrenCounts $child
                $totalChildren += $childCounts.Total
                $checkedChildren += $childCounts.Checked
            }
        }
        return @{
            Total = $totalChildren
            Checked = $checkedChildren
        }
    }
    
    function Update-NodeTextWithChildCounts {
        param($node)
        if ($node.Nodes.Count -gt 0) {
            $counts = Get-ChildrenCounts $node
            # Remove any existing count suffix
            $baseText = $node.Text
            $countPattern = " \(\d+/\d+\)$"
            if ($baseText -match $countPattern) { $baseText = $baseText -replace $countPattern, "" }
            # Update the text with the new count
            $node.Text = "$baseText ($($counts.Checked)/$($counts.Total))"
        }
    }
    
    function Update-AllNodeTexts {
        param($nodes)
        foreach ($node in $nodes) {
            if ($node.Nodes.Count -gt 0) {
                # Update children first (depth-first)
                Update-AllNodeTexts $node.Nodes
                # Then update this node
                Update-NodeTextWithChildCounts $node
            }
        }
    }
    
    New-Button "Select All" 10 20 100 25 { Update-NodesChecked $tree.Nodes $true }
    New-Button "Unselect All" 110 20 100 25 { Update-NodesChecked $tree.Nodes $false }
    
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.CheckBoxes = $true
    $tree.Location = New-Object System.Drawing.Point(10, 60)
    $tree.Size = New-Object System.Drawing.Size($($dlg.Width - 35), $($dlg.Height - 110))
    $tree.Anchor = "Top,Left,Bottom,Right"

    function Update-ChildNodesChecked ($node, $state) {
        foreach ($child in $node.Nodes) {
            $child.Checked = $state
            if ($child.Nodes.Count -gt 0) { Update-ChildNodesChecked $child $state }
        }
    }
    
    function Expand-CheckedParents ($nodes) {
        foreach ($node in $nodes) {
            if ($node.Nodes.Count -gt 0) {
                $result = Get-ChildrenCounts -node $node
                if ($result.Checked -gt 0 -and $result.Checked -lt $result.Total) { $node.Expand() }
                # Recursive call to process sub-nodes
                Expand-CheckedParents $node.Nodes
            }
        }
    }

    function Get-StoredBookmark ($node, $storedSelection) {
        if (-not $storedSelection -or $storedSelection.Count -eq 0) { return $null }
        $decoded = if ([string]::IsNullOrWhiteSpace($node.name)) { $node.url } else { [System.Web.HttpUtility]::HtmlDecode($node.name) }
        foreach ($stored in $storedSelection) {
            # Remove counters from stored text before comparison
            $storedText = $stored.Text -replace " \(\d+/\d+\)$", ""
            if (($node.type -eq 'url' -and $stored.Type -eq "Bookmark" -and $decoded -eq $storedText -and $stored.Tag -eq $node.url) -or 
                ($node.children -and $stored.Type -eq "Folder" -and $decoded -eq $storedText)) {
                return $stored
            }
        }
        return $null
    }

    $key = if ($mode -eq "restore" -and $null -eq $checkbox) { "Common|Common" } else { "$browser|$BrowserProfile" }

    if ($mode -eq "restore") {
        $storedPaths = @()
        if ($script:RestoreSelectedUrls.ContainsKey($key)) {
            $storedPaths = $script:RestoreSelectedUrls[$key] | ForEach-Object { $_.Path }
        }
        function Add-RestoreNodes ($path, $parent) {
            Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $tn = New-Object System.Windows.Forms.TreeNode($_.Name)
                $tn.Tag = "folder"
                if ($null -eq $parent) { $tree.Nodes.Add($tn) } else { $parent.Nodes.Add($tn) }
                Add-RestoreNodes -path $_.FullName -parent $tn
            }
            Get-ChildItem -Path $path -Filter "*.url" -File -ErrorAction SilentlyContinue | ForEach-Object {
                $tn = New-Object System.Windows.Forms.TreeNode(([System.IO.Path]::GetFileNameWithoutExtension($_.Name)))
                $tn.Tag = $_.FullName
                $tn.Checked = if ($storedPaths.Count -gt 0) { ($storedPaths -contains $_.FullName) } else { $true }
                if ($null -eq $parent) { $tree.Nodes.Add($tn) } else { $parent.Nodes.Add($tn) }
            }
        }
        Add-RestoreNodes -path $source -parent $null

    } else {
        if (-not (Test-Path $source)) {
            Log "Bookmarks file not found: $source"
            [System.Windows.Forms.MessageBox]::Show("Bookmarks file not found: $source", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $json = Get-Content -Path $source -Raw -Encoding UTF8 | ConvertFrom-Json
        $storedSelection = @(); $forceState = $true
        if ($script:BackupSelectedBookmarks.ContainsKey($key)) {
            $storedSelection = $script:BackupSelectedBookmarks[$key]; $forceState = $null
        }
        function Add-BookmarksToTreeNode ($node, $parentCollection, $storedSelection, $forceState) {
            if ($node.type -eq 'url') {
                $name = if ([string]::IsNullOrWhiteSpace($node.name)) { $node.url } else { [System.Web.HttpUtility]::HtmlDecode($node.name) }
                $tn = $parentCollection.Add($name)
                $tn.Tag = $node.url
                $tn.Checked = if ($null -ne $forceState) { $forceState } else {
                    $stored = Get-StoredBookmark $node $storedSelection
                    ($null -ne $stored -and $stored.Checked)
                }
            }
            if ($node.children) {
                $name = if ([string]::IsNullOrWhiteSpace($node.name)) { "Folder" } else { [System.Web.HttpUtility]::HtmlDecode($node.name) }
                $tn = $parentCollection.Add($name)
                if ($null -ne $forceState) {
                    $tn.Checked = $forceState
                    $childForce = $forceState; $childStored = @()
                } else {
                    $storedFolder = $storedSelection | Where-Object { $_.Type -eq "Folder" -and ([System.Web.HttpUtility]::HtmlDecode($node.name)) -eq $_.Text }
                    if ($storedFolder) {
                        $tn.Checked = $storedFolder.Checked
                        $childForce = $null; $childStored = $storedFolder.Children
                    } else {
                        $tn.Checked = $false
                        $childForce = $false; $childStored = @()
                    }
                }
                foreach ($child in $node.children) { Add-BookmarksToTreeNode $child $tn.Nodes $childStored $childForce }
            }
        }
        foreach ($child in $json.roots.bookmark_bar.children) { Add-BookmarksToTreeNode $child $tree.Nodes $storedSelection $forceState }
    }

    function Update-ParentStates ($nodes) {
        foreach ($n in $nodes) {
            if ($n.Nodes.Count -gt 0) {
                Update-ParentStates $n.Nodes
                # Checks if at least one child node is checked
                $n.Checked = $null -ne ($n.Nodes | Where-Object { $_.Checked })
            }
        }
    }

    Update-ParentStates $tree.Nodes
    Expand-CheckedParents $tree.Nodes
    # Initialize node texts with child counters (checked_nodes_number/total_nodes_number)
    Update-AllNodeTexts $tree.Nodes

    New-Button "Collapse All" ($dlg.ClientSize.Width - 100) 20 90 25 { $tree.CollapseAll() }
    New-Button "Expand All" ($dlg.ClientSize.Width - 190) 20 90 25 { $tree.ExpandAll() }

    $tree.Add_AfterCheck({
        param($s, $e)
        if ($script:ignoreAfterCheck) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $script:ignoreAfterCheck = $true
        try {
            if ($e.Action -ne [System.Windows.Forms.TreeViewAction]::Unknown) { Update-ChildNodesChecked $e.Node $e.Node.Checked }
            $current = $e.Node.Parent
            while ($null -ne $current) {
                # checks if at least one child node is checked
                $current.Checked = $null -ne ($current.Nodes | Where-Object { $_.Checked })
                $current = $current.Parent
            }
            Update-AllNodeTexts $tree.Nodes
        } finally { $script:ignoreAfterCheck = $false ; $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor }
    })
    $dlg.Controls.Add($tree)

    $btnOK = New-Button "OK" ($dlg.Width / 2 - 50) 10 100 40 {
        if ($mode -eq "restore") {
            function Get-CheckedFiles ($nodes) {
                $result = @()
                foreach ($n in $nodes) {
                    if ($n.Checked -and $n.Tag -and $n.Tag -ne "folder") { $result += $n }
                    if ($n.Nodes.Count -gt 0) { $result += Get-CheckedFiles $n.Nodes }
                }
                return $result
            }
            $urls = @()
            foreach ($n in (Get-CheckedFiles $tree.Nodes)) {
                $fp = $n.Tag
                if (Test-Path $fp) {
                    $content = Get-Content -Path $fp -Encoding UTF8 -ErrorAction SilentlyContinue
                    if ($content -and ($urlLine = $content | Where-Object { $_ -like "URL=*" })) { $urls += @{ Name = $n.Text; URL = $urlLine.Substring(4).Trim(); Path = $fp } }
                }
            }
            $script:RestoreSelectedUrls[$key] = $urls
            if ($null -ne $checkbox) { $checkbox.Text = "$BrowserProfile ($($urls.Count) URL files)" }
            Log "Selected $($urls.Count) URL files for restore mode: $key"
        } else {
            function Get-CheckedNodes ($nodes) {
                $results = @()
                $count = 0
                foreach ($n in $nodes) {
                    # Remove counters from node text before storing
                    $baseText = $n.Text -replace " \(\d+/\d+\)$", ""
                    if ($n.Nodes.Count -eq 0) {
                        if ($n.Checked) {
                            $results += @{ Text = $baseText; Tag = $n.Tag; Type = "Bookmark"; Checked = $true }
                            $count++
                        }
                    } else {
                        $childResult = Get-CheckedNodes $n.Nodes
                        $results += @{ Text = $baseText; Type = "Folder"; Children = $childResult.Children; Checked = $n.Checked }
                        $count += $childResult.BookmarkCount
                    }
                }
                return @{ Children = $results; BookmarkCount = $count }
            }
            $result = Get-CheckedNodes $tree.Nodes
            $script:BackupSelectedBookmarks[$key] = $result.Children
            $total = Get-BookmarksCount $source
            $checkbox.Text = "$BrowserProfile ($($result.BookmarkCount) / $total)"
            $checkbox.Checked = $true
            Log "Selected $($result.BookmarkCount) bookmarks for backup mode: $key"
        }
        $dlg.Close()
    }
    $btnOK.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $btnOK.Anchor = "Top,Right,Left"
    
    $dlg.ShowDialog()
}

function Show-ProgressBar {
    param([string]$title, [int]$maxValue)
    Log "Showing progress bar: $title, max=$maxValue"
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(300,100)
    $form.StartPosition = "CenterScreen"
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(10,20)
    $bar.Size = New-Object System.Drawing.Size(260,20)
    $bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $bar.Minimum = 0
    $bar.Maximum = [Math]::Max(1, $maxValue)
    $bar.Value = 0
    $form.Controls.Add($bar)
    $form.Show()
    return @{ Form = $form; Bar = $bar }
}

function Export-Bookmarks {
    Log "Starting Export-Bookmarks function"
    function Get-RecursiveBookmarksCount ($nodes) {
        $count = 0
        foreach ($node in $nodes) {
            if ($node.Type -eq "Bookmark" -and $node.Checked) { $count++ }
            elseif ($node.Type -eq "Folder" -and $node.Checked) { $count += Get-RecursiveBookmarksCount $node.Children }
        }
        return $count
    }
    function Convert-SelectedBookmarks {
        param([array]$nodes, [string]$currentPath, [ref]$currentBookmark, $progress)
        $nodesToProcess = $nodes | Where-Object { $_.Checked -eq $true }
        function Get-ValidName($name, $url, $isFolder = $false) {
            if (-not $isFolder -and $name -match '^(https?://)?(www\.)?') { $name = $name -replace '^(https?://)?(www\.)?', '' }
            $valid = $name -replace '[^\p{L}\p{Nd}\s@''._-]', '_' -replace '[\x00-\x1F]', '' -replace '_+', '_'
            $valid = $valid.Trim(' ._-')
            if (-not $isFolder -and ([string]::IsNullOrWhiteSpace($valid) -or $valid -match '^[\s._-]+$')) {
                $valid = $url -replace '^(https?://)?(www\.)?', '' -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
                $valid = $valid.Trim(' ._-')
                if ([string]::IsNullOrWhiteSpace($valid)) { $valid = "Unnamed_Bookmark" }
            }
            if ($valid.Length -gt 60) { $valid = $valid.Substring(0,60).TrimEnd(' ._-') }
            return $valid
        }
        foreach ($node in $nodesToProcess) {
            if ($node.Type -eq "Bookmark") {
                $url = $node.Tag
                if ($url -match '^(chrome|edge)://') {
                    Log "Skipping browser-specific URL: $url"
                    continue
                }
                $validName = Get-ValidName $node.Text $url
                $filePath = Join-Path -Path $currentPath -ChildPath "$validName.url"
                if (-not (Test-Path -LiteralPath $filePath)) {
                    try {
                        @"
[InternetShortcut]
URL=$url
"@ | Out-File -LiteralPath $filePath -Encoding UTF8 -Force -ErrorAction Stop
                    } catch {
                        Log "Failed to create URL file: $filePath. Error: $($_.Exception.Message)"
                        $script:ExportErrors += "Failed to create file: $filePath"
                    }
                } else { Log "File already exists: $filePath. Skipping creation to avoid duplicates." }
                $currentBookmark.Value++
                $progress.Bar.Value = $currentBookmark.Value
            } elseif ($node.Type -eq "Folder") {
                $folderName = Get-ValidName $node.Text '' $true
                $folderPath = Join-Path -Path $currentPath -ChildPath $folderName
                try {
                    if (-not (Test-Path -LiteralPath $folderPath)) { New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null }
                    Convert-SelectedBookmarks $node.Children $folderPath $currentBookmark $progress
                } catch {
                    Log "Failed to create folder: $folderPath. Error: $($_.Exception.Message)"
                    $script:ExportErrors += "Failed to create folder: $folderPath"
                }
            }
        }
    }
    $selectedProfiles = @()
    $script:backupChromeCheckboxes | Where-Object { $_.Checked } | ForEach-Object { 
        $selectedProfiles += @{ Browser = "Chrome"; Profile = $_.Tag.Profile } 
    }
    $script:backupEdgeCheckboxes | Where-Object { $_.Checked } | ForEach-Object { 
        $selectedProfiles += @{ Browser = "Edge"; Profile = $_.Tag.Profile } 
    }
    if ($selectedProfiles.Count -eq 0) {
        Log "No profiles selected for export"
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Browser Profile to export.", 
            "No Profile Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $exportFolder = $script:targetPath
    Log "Export target folder: $exportFolder"
    
    $totalBookmarks = ($selectedProfiles | ForEach-Object {
        $key = "$($_.Browser)|$($_.Profile)"
        if ($script:BackupSelectedBookmarks.ContainsKey($key)) {
            $count = Get-RecursiveBookmarksCount $script:BackupSelectedBookmarks[$key]
            $count
        } else { 0 }
    } | Measure-Object -Sum).Sum
    
    Log "Total bookmarks to export: $totalBookmarks"
    $totalBookmarks = [Math]::Max(1, $totalBookmarks)
    $progress = Show-ProgressBar "Exporting Bookmarks" $totalBookmarks
    $currentBookmark = 0
    
    foreach ($profile in $selectedProfiles) {
        Log "Creating URL files... browser: $($profile.Browser), profile: $($profile.Profile)"
        $key = "$($profile.Browser)|$($profile.Profile)"
        if ($script:BackupSelectedBookmarks.ContainsKey($key)) { Convert-SelectedBookmarks $script:BackupSelectedBookmarks[$key] $exportFolder ([ref]$currentBookmark) $progress } 
        else { Log "No bookmarks found for profile $key" }
        Log "URL files created. browser: $($profile.Browser), profile: $($profile.Profile)"
    }
    
    Reset-UI
    $progress.Form.Close()
    
    if ($script:ExportErrors.Count -gt 0) {
        Log "Export completed with $($script:ExportErrors.Count) errors"
        [System.Windows.Forms.MessageBox]::Show("Some errors occurred during export.", "Errors", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        Log "Export completed successfully"
        [System.Windows.Forms.MessageBox]::Show("Bookmarks exported successfully.", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    $script:ExportErrors = @()
}

function Import-bookmarks {
    Log "Starting Import-bookmarks function"
    function Get-ChromeTimestamp {
        return [math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date "1601-01-01")).TotalMicroseconds)
    }
    function Test-BookmarksFile ($bookmarksFile) {
        if (-not (Test-Path $bookmarksFile)) {
            Log "Creating new bookmarks file: $bookmarksFile"
            $initial = @{
                roots = @{
                    bookmark_bar = @{ children = @(); type = "folder" }
                    other = @{ children = @(); type = "folder" }
                    synced = @{ children = @(); type = "folder" }
                }
                version = 1
            }
            $initial | ConvertTo-Json -Depth 100 | Set-Content -Path $bookmarksFile -Encoding UTF8
        }
    }
    function Test-BookmarkFolder {
        param([object]$ParentNode, [string]$FolderName)
        $existing = $ParentNode.children | Where-Object { $_.name -eq $FolderName -and $_.type -eq "folder" }
        if ($existing) {
            return $existing
        } else {
            $timestamp = Get-ChromeTimestamp
            $newFolder = [PSCustomObject]@{
                date_added     = $timestamp.ToString()
                date_last_used = "0"
                date_modified  = $timestamp.ToString()
                guid           = "new-guid-" + ([guid]::NewGuid().ToString().Substring(0,8))
                id             = $timestamp.ToString()
                name           = $FolderName
                source         = "user_add"
                type           = "folder"
                children       = @()
            }
            $ParentNode.children += $newFolder
            return $newFolder
        }
    }
    function Get-RelativePath {
        param([string]$FullPath, [string]$BasePath)
        if (-not $BasePath.EndsWith("\")) { $BasePath += "\" }
        $uriFull = New-Object System.Uri($FullPath)
        $uriBase = New-Object System.Uri($BasePath)
        $relativeUri = $uriBase.MakeRelativeUri($uriFull).ToString()
        return [System.Uri]::UnescapeDataString($relativeUri).Replace("/", "\")
    }
    function Import-SelectedUrls {
        param([object]$bookmarksData, [array]$urlObjects, [string]$baseFolder)
        Log "Importing URLs using base folder: $baseFolder"
        foreach ($obj in $urlObjects) {
            # Determine .url file directory
            $fileDirectory = [System.IO.Path]::GetDirectoryName($obj.Path)
            $relativePath = ""
            if ($fileDirectory -and $baseFolder) { $relativePath = Get-RelativePath -FullPath $fileDirectory -BasePath $baseFolder }
            # Start from 'bookmark_bar' node
            $currentNode = $bookmarksData.roots.bookmark_bar
            if ($relativePath -ne "") {
                $folders = $relativePath -split "\\"
                foreach ($folder in $folders) {
                    if ($folder -ne "") { $currentNode = Test-BookmarkFolder -ParentNode $currentNode -FolderName $folder }
                }
            }
            # Add URL to current folder if it doesn't exist already
            $exists = $currentNode.children | Where-Object { $_.type -eq "url" -and $_.url -eq $obj.URL }
            if (-not $exists) {
                $timestamp = Get-ChromeTimestamp
                $newBookmark = [PSCustomObject]@{
                    date_added     = $timestamp.ToString()
                    date_last_used = "0"
                    guid           = "new-guid-" + ([guid]::NewGuid().ToString().Substring(0,8))
                    id             = $timestamp.ToString()
                    name           = $obj.Name
                    show_icon      = $false
                    source         = "user_add"
                    type           = "url"
                    url            = $obj.URL
                }
                $currentNode.children += $newBookmark
            }
        }
        return $bookmarksData
    }
    function Test-BrowserWindowOpen {
        param ([string]$browserName)
        $processes = Get-Process -Name $browserName -ErrorAction SilentlyContinue
        foreach ($process in $processes) {
            if ($process.MainWindowHandle -ne [IntPtr]::Zero -and $process.MainWindowTitle -ne "") {
                return $true
            }
        }
        return $false
    }

    $selected = @()
    foreach ($chk in $script:restoreChromeCheckboxes) { if ($chk.Checked) { $selected += $chk.Tag } }
    foreach ($chk in $script:restoreEdgeCheckboxes) { if ($chk.Checked) { $selected += $chk.Tag } }
    if ($selected.Count -eq 0) {
        Log "No profiles selected for import"
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Browser Profile to import to.", "No Profile Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $chromeRunningBefore = $selected | Where-Object { $_.Browser -eq "Chrome" } | ForEach-Object { Test-BrowserWindowOpen -browserName "chrome" }
    $edgeRunningBefore = $selected | Where-Object { $_.Browser -eq "Edge" } | ForEach-Object { Test-BrowserWindowOpen -browserName "msedge" }

    if (($chromeRunningBefore -or $edgeRunningBefore) -and (-not ($script:autorestore))) {
        Log "Browsers are running - asking user for confirmation"
        $res = [System.Windows.Forms.MessageBox]::Show("Selected browsers will be closed and restarted. Continue?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::OKCancel)
        if ($res -ne [System.Windows.Forms.DialogResult]::OK) { 
            Log "User canceled import operation"
            return 
        }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $progForm = Show-ProgressBar -title "Importing Bookmarks" -maxValue 100
    function Update-Progress ($val) {
        $progForm.Bar.Value = $val
        $progForm.Form.Refresh()
    }
    Update-Progress 10
    if ($selected | Where-Object { $_.Browser -eq "Chrome" }) { 
        Log "Closing Chrome browser"
        taskkill /f /im chrome.exe | Out-Null 
    }
    if ($selected | Where-Object { $_.Browser -eq "Edge" }) { 
        Log "Closing Edge browser"
        taskkill /f /im msedge.exe | Out-Null 
    }
    Start-Sleep -Seconds 1
    Update-Progress 30
    
    foreach ($BrowserProfile in $selected) {
        $browser = $BrowserProfile.Browser
        $profName = $BrowserProfile.Profile
        $bmFile = if ($browser -eq "Chrome") { "$env:LOCALAPPDATA\Google\Chrome\User Data\$profName\Bookmarks" }
                else { "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$profName\Bookmarks" }
        
        Log "Importing bookmarks to $browser profile $profName"
        Test-BookmarksFile $bmFile
        $bmData = Get-Content -Path $bmFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $bmData = Import-SelectedUrls -bookmarksData $bmData -urlObjects $script:CommonRestoreUrls -baseFolder $script:sourcePath
        $jsonOut = $bmData | ConvertTo-Json -Depth 100 -Compress
        [System.IO.File]::WriteAllText($bmFile, $jsonOut, [System.Text.Encoding]::UTF8)
        Log "Finished importing to $browser profile $profName"
    }
    Update-Progress 50
    Reset-UI
    Update-Progress 90

    if ($chromeRunningBefore) { 
        Log "Restarting Chrome browser"
        Start-Process "chrome.exe"
    }
    if ($edgeRunningBefore) { 
        Log "Restarting Edge browser"
        Start-Process "msedge.exe"
    }

    Update-Progress 100
    $progForm.Form.Close()
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor 
    Log "Import completed successfully"
    if (-not ($script:autorestore)) {
        [System.Windows.Forms.MessageBox]::Show("Import completed successfully!", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}

function Reset-UI {
    Log "Resetting UI"
    $script:BackupSelectedBookmarks = @{}
    $script:RestoreSelectedUrls = @{}    
    $script:CommonRestoreUrls = @{}      
    $script:ExportErrors = @()

    $backupChromeGroup.Controls.Clear()
    $backupEdgeGroup.Controls.Clear()
    $chromeProfiles = Get-BrowserProfiles "chrome"
    $edgeProfiles = Get-BrowserProfiles "edge"
    $script:backupChromeCheckboxes = New-ProfileCheckboxes -group $backupChromeGroup -BrowserProfiles $chromeProfiles -browser "Chrome" -mode "backup"
    $script:backupEdgeCheckboxes   = New-ProfileCheckboxes -group $backupEdgeGroup -BrowserProfiles $edgeProfiles -browser "Edge" -mode "backup"
    
    $restoreChromeGroup.Controls.Clear()
    $restoreEdgeGroup.Controls.Clear()
    $script:restoreChromeCheckboxes = New-ProfileCheckboxes -group $restoreChromeGroup -BrowserProfiles $chromeProfiles -browser "Chrome" -mode "restore"
    $script:restoreEdgeCheckboxes   = New-ProfileCheckboxes -group $restoreEdgeGroup -BrowserProfiles $edgeProfiles -browser "Edge" -mode "restore"
    
    # Use source path from dropdown
    $urlFiles = Get-ChildItem -Path $script:sourcePath -Filter "*.url" -Recurse -ErrorAction SilentlyContinue
    $commonCount = $urlFiles.Count
    Log "Found $commonCount URL files in source folder: $script:sourcePath"
    
    # CommonRestoreUrls initialization
    $urls = @()
    foreach ($file in $urlFiles) {
        $content = Get-Content -Path $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($content -and ($urlLine = $content | Where-Object { $_ -like "URL=*" })) {
            $urls += @{ 
                Name = [System.IO.Path]::GetFileNameWithoutExtension($file.Name); 
                URL = $urlLine.Substring(4).Trim(); 
                Path = $file.FullName 
            }
        }
    }
    $script:CommonRestoreUrls = $urls
    $script:RestoreSelectedUrls["Common|Common"] = $urls
    $commonSelectButton.Text = "Select URLs to import ($commonCount)"
}

##############################################
# MAIN INTERFACE
##############################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fast Bookmarks Manager"
$form.Size = New-Object System.Drawing.Size(605,550)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial",10)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$tabControl.Location = New-Object System.Drawing.Point(0, 60)
$tabControl.Size = New-Object System.Drawing.Size($form.ClientSize.Width, ($form.ClientSize.Height - 60))

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Size = New-Object System.Drawing.Size(80,25)
$refreshButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - 90), 0)
$refreshButton.Anchor = "Top, Right"
$refreshButton.Add_Click({ 
    Log "Refresh button clicked"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    Reset-UI
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor 
})
$form.Controls.Add($refreshButton)

$tabBackup = New-Object System.Windows.Forms.TabPage
$tabBackup.Text = "Backup"
$backupChromeGroup = New-Object System.Windows.Forms.GroupBox
$backupChromeGroup.Location = New-Object System.Drawing.Point(20,20)
$backupChromeGroup.Size = New-Object System.Drawing.Size(260,380)
$backupChromeGroup.Text = "Chrome Profiles"
$tabBackup.Controls.Add($backupChromeGroup)
$backupEdgeGroup = New-Object System.Windows.Forms.GroupBox
$backupEdgeGroup.Location = New-Object System.Drawing.Point(300,20)
$backupEdgeGroup.Size = New-Object System.Drawing.Size(260,380)
$backupEdgeGroup.Text = "Edge Profiles"
$tabBackup.Controls.Add($backupEdgeGroup)

$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Text = "Settings"
$settingsGroup = New-Object System.Windows.Forms.GroupBox
$settingsGroup.Location = New-Object System.Drawing.Point(20, 20)
$settingsGroup.Size = New-Object System.Drawing.Size(540, 380)
$settingsGroup.Text = "Folder Settings"
$tabSettings.Controls.Add($settingsGroup)

$targetGroupBox = New-Object System.Windows.Forms.GroupBox
$targetGroupBox.Location = New-Object System.Drawing.Point(20, 30)
$targetGroupBox.Size = New-Object System.Drawing.Size(500, 100)
$targetGroupBox.Text = "Backup Target Folder"
$settingsGroup.Controls.Add($targetGroupBox)
$targetFolderLabel = New-Object System.Windows.Forms.Label
$targetFolderLabel.Text = "Select where to save backups:"
$targetFolderLabel.Size = New-Object System.Drawing.Size(200, 20)
$targetFolderLabel.Location = New-Object System.Drawing.Point(10, 30)
$targetGroupBox.Controls.Add($targetFolderLabel)
$targetFolderTextBox = New-Object System.Windows.Forms.TextBox
$targetFolderTextBox.Size = New-Object System.Drawing.Size(300, 25)
$targetFolderTextBox.Location = New-Object System.Drawing.Point(10, 55)
$targetFolderTextBox.Text = (Get-Location).Path
$targetGroupBox.Controls.Add($targetFolderTextBox)
$browseTargetButton = New-Object System.Windows.Forms.Button
$browseTargetButton.Text = "Browse..."
$browseTargetButton.Size = New-Object System.Drawing.Size(80, 25)
$browseTargetButton.Location = New-Object System.Drawing.Point(320, 55)
$browseTargetButton.Add_Click({
    $selectedPath = ([FolderSelectDialog]::new($targetFolderTextBox.Text, "Select Backup Target Folder", "Please select the folder to store backups.")).getPath()
    if ($selectedPath -ne "") { $targetFolderTextBox.Text = $selectedPath }
})
$targetGroupBox.Controls.Add($browseTargetButton)

$sourceGroupBox = New-Object System.Windows.Forms.GroupBox
$sourceGroupBox.Location = New-Object System.Drawing.Point(20, 150)
$sourceGroupBox.Size = New-Object System.Drawing.Size(500, 100)
$sourceGroupBox.Text = "Restore Source Folder"
$settingsGroup.Controls.Add($sourceGroupBox)
$sourceFolderLabel = New-Object System.Windows.Forms.Label
$sourceFolderLabel.Text = "Select where to restore backups from:"
$sourceFolderLabel.Size = New-Object System.Drawing.Size(250, 20)
$sourceFolderLabel.Location = New-Object System.Drawing.Point(10, 30)
$sourceGroupBox.Controls.Add($sourceFolderLabel)
$sourceFolderTextBox = New-Object System.Windows.Forms.TextBox
$sourceFolderTextBox.Size = New-Object System.Drawing.Size(300, 25)
$sourceFolderTextBox.Location = New-Object System.Drawing.Point(10, 55)
$sourceFolderTextBox.Text = (Get-Location).Path
$sourceFolderTextBox.ReadOnly = $true
$sourceGroupBox.Controls.Add($sourceFolderTextBox)

$browseSourceButton = New-Object System.Windows.Forms.Button
$browseSourceButton.Text = "Browse..."
$browseSourceButton.Size = New-Object System.Drawing.Size(80, 25)
$browseSourceButton.Location = New-Object System.Drawing.Point(320, 55)
$browseSourceButton.Add_Click({
    $selectedPath = ([FolderSelectDialog]::new($sourceFolderTextBox.Text, "Select Restore Source Folder", "Please select the folder to restore backups from.")).getPath()
    if ($selectedPath -ne "") { $sourceFolderTextBox.Text = $selectedPath }
})
$sourceGroupBox.Controls.Add($browseSourceButton)

$targetFolderTextBox.Add_TextChanged({
    $script:targetPath = $targetFolderTextBox.Text
    Log "Backup target folder changed: $script:targetPath"
})
$sourceFolderTextBox.Add_TextChanged({
    $script:sourcePath = $sourceFolderTextBox.Text
    Log "Restore source folder changed: $script:sourcePath"
    Reset-UI
})

$script:sourcePath = (Get-Location).Path
$script:targetPath = (Get-Location).Path

function New-ProfileCheckboxes {
    param([System.Windows.Forms.GroupBox]$group, [array]$BrowserProfiles, [string]$browser, [ValidateSet("backup", "restore")][string]$mode)
    function Get-AllBookmarks($f) {
        $json = Get-Content -Path $f -Raw -Encoding UTF8 | ConvertFrom-Json
        function Get-BM($nodes) {
            $r = @()
            foreach ($n in $nodes) {
                if ($n.type -eq 'url') {
                    $nm = if ([string]::IsNullOrWhiteSpace($n.name)) {$n.url} else {[System.Web.HttpUtility]::HtmlDecode($n.name)}
                    $r += @{Text=$nm; Tag=$n.url; Type="Bookmark"; Checked=$true}
                } elseif ($n.children) {
                    $nm = if ([string]::IsNullOrWhiteSpace($n.name)) {"Folder"} else {[System.Web.HttpUtility]::HtmlDecode($n.name)}
                    $c = Get-BM $n.children
                    if ($c.Count -gt 0) {$r += @{Text=$nm; Type="Folder"; Children=$c; Checked=$true}}
                }
            }
            return $r
        }
        return Get-BM $json.roots.bookmark_bar.children
    }
    $y = 30
    $chkList = @()
    if (-not $BrowserProfiles) {
        $l = New-Object System.Windows.Forms.Label
        $l.Location = New-Object System.Drawing.Point(10, $y)
        $l.Size = New-Object System.Drawing.Size(240,20)
        $l.Text = "$browser not found"
        $group.Controls.Add($l)
        return $chkList
    }
    $btnText = if ($mode -eq "backup") {"Select"} else {"View"}
    foreach ($p in $BrowserProfiles) {
        $bmFile = if ($browser -eq "Chrome") {"$env:LOCALAPPDATA\Google\Chrome\User Data\$p\Bookmarks"} 
                  else {"$env:LOCALAPPDATA\Microsoft\Edge\User Data\$p\Bookmarks"}
        $bmExists = Test-Path $bmFile
        $bmCount = if ($bmExists) {Get-BookmarksCount $bmFile} else {0}
        if (!$bmExists) {$sfx = " (File not found)"; $enableBtn = $false}
        elseif ($bmCount -eq 0) {$sfx = " (No bookmarks)"; $enableBtn = $false}
        else {
            $sfx = if ($mode -eq "backup") {" ($bmCount / $bmCount)"} else {" ($bmCount)"}
            $enableBtn = $true
            if ($mode -eq "backup") {$script:BackupSelectedBookmarks["$browser|$p"] = Get-AllBookmarks $bmFile}
        }
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Location = New-Object System.Drawing.Point(10, $y)
        $chk.Size = New-Object System.Drawing.Size(180,20)
        $chk.Text = $p + $sfx
        $chk.Enabled = if ($mode -eq "backup") {$bmExists -and $bmCount -gt 0} else {$true}
        $chk.Tag = @{Browser=$browser; Profile=$p}
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point(190, $y)
        $btn.Size = New-Object System.Drawing.Size(60,20)
        $btn.Text = $btnText
        $btn.Enabled = $enableBtn
        $btn.Tag = @{Browser=$browser; Profile=$p; BookmarksFile=$bmFile; Checkbox=$chk}
        $btn.Add_Click({
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            [System.Windows.Forms.Application]::DoEvents()
            $t = $this.Tag
            if ($this.Text -eq "Select") { Show-TreeView -mode "backup" -browser $t.Browser -BrowserProfile $t.Profile -source $t.BookmarksFile -checkbox $t.Checkbox }
            else { Show-SimpleBookmarksTreeView $t.Browser $t.Profile $t.BookmarksFile }
            $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor
        })
        $group.Controls.Add($chk)
        $group.Controls.Add($btn)
        $chkList += $chk
        $y += 30
    }
    return $chkList
}

$chromeProfiles = Get-BrowserProfiles "chrome"
$edgeProfiles = Get-BrowserProfiles "edge"
$script:backupChromeCheckboxes = New-ProfileCheckboxes -group $backupChromeGroup -BrowserProfiles $chromeProfiles -browser "Chrome" -mode "backup"
$script:backupEdgeCheckboxes   = New-ProfileCheckboxes -group $backupEdgeGroup -BrowserProfiles $edgeProfiles -browser "Edge" -mode "backup"

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Size = New-Object System.Drawing.Size(100,30)
$exportButton.Location = New-Object System.Drawing.Point($($form.ClientSize.Width/2 - 50), 440)
$exportButton.Text = "Export"
$exportButton.Add_Click({ Export-Bookmarks })
$tabBackup.Controls.Add($exportButton)

$tabRestore = New-Object System.Windows.Forms.TabPage
$tabRestore.Text = "Restore"

$commonSelectButton = New-Object System.Windows.Forms.Button
$commonSelectButton.Size = New-Object System.Drawing.Size(200,30)
$commonSelectButton.Text = ""
$commonSelectButton.Location = New-Object System.Drawing.Point($($form.ClientSize.Width/2 - 100), 10)
$commonSelectButton.Text = "Select URLs to import (0)"
$commonSelectButton.Add_Click({
    Log "Common select button clicked"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    Show-TreeView -mode "restore" -browser "Common" -BrowserProfile "Common" -source $script:sourcePath -checkbox $null
    $form.Cursor = [System.Windows.Forms.Cursors]::DefaultCursor
    if ($script:RestoreSelectedUrls.ContainsKey("Common|Common")) {
         $cnt = ($script:RestoreSelectedUrls["Common|Common"]).Count
         $commonSelectButton.Text = "Select URLs to import ($cnt)"
         $script:CommonRestoreUrls = $script:RestoreSelectedUrls["Common|Common"]
         Log "Selected $cnt URLs for common import"
    } else {
         $commonSelectButton.Text = "Select URLs to import (0)"
         $script:CommonRestoreUrls = @()
    }
})
$tabRestore.Controls.Add($commonSelectButton)

$restoreChromeGroup = New-Object System.Windows.Forms.GroupBox
$restoreChromeGroup.Location = New-Object System.Drawing.Point(20,50)
$restoreChromeGroup.Size = New-Object System.Drawing.Size(260,380)
$restoreChromeGroup.Text = "Chrome Profiles"
$tabRestore.Controls.Add($restoreChromeGroup)
$restoreEdgeGroup = New-Object System.Windows.Forms.GroupBox
$restoreEdgeGroup.Location = New-Object System.Drawing.Point(300,50)
$restoreEdgeGroup.Size = New-Object System.Drawing.Size(260,380)
$restoreEdgeGroup.Text = "Edge Profiles"
$tabRestore.Controls.Add($restoreEdgeGroup)

$script:restoreChromeCheckboxes = New-ProfileCheckboxes -group $restoreChromeGroup -BrowserProfiles $chromeProfiles -browser "Chrome" -mode "restore"
$script:restoreEdgeCheckboxes   = New-ProfileCheckboxes -group $restoreEdgeGroup -BrowserProfiles $edgeProfiles -browser "Edge" -mode "restore"

$importButton = New-Object System.Windows.Forms.Button
$importButton.Size = New-Object System.Drawing.Size(100,30)
$importButton.Location = New-Object System.Drawing.Point($($form.ClientSize.Width/2 - 50), 440)
$importButton.Text = "Import"
$importButton.Add_Click({ Import-bookmarks })
$tabRestore.Controls.Add($importButton)

$tabControl.TabPages.Add($tabBackup)
$tabControl.TabPages.Add($tabRestore)
$tabControl.TabPages.Add($tabSettings)
$form.Controls.Add($tabControl)

Reset-UI

if ($script:autorestore) {
    Log "Auto-restore mode enabled"
    $form.Add_Shown({
        # Use BeginInvoke to allow UI to render fully before performing auto actions
        $form.BeginInvoke([Action]{
            Log "Starting auto-restore sequence"
            $tabControl.SelectedTab = $tabRestore
            $defaultsFound = $false
            foreach ($chk in $script:restoreChromeCheckboxes) {
                if ($chk.Enabled -and $chk.Text.StartsWith("Default")) {
                    Log "Checking Chrome Default profile: $($chk.Text)"
                    $chk.Checked = $true
                    $defaultsFound = $true
                }
            }
            foreach ($chk in $script:restoreEdgeCheckboxes) {
                if ($chk.Enabled -and $chk.Text.StartsWith("Default")) {
                    Log "Checking Edge Default profile: $($chk.Text)"
                    $chk.Checked = $true
                    $defaultsFound = $true
                }
            }
            if (-not $defaultsFound) {
                Log "No Default profiles found or they were disabled"
                $form.Close()
                return
            }
            Log "Triggering import process"
            Import-bookmarks
            Log "Auto-restore completed, closing application"
            $form.Close()
        })
    })
}

$form.ShowDialog()
Log "--- Application closed ---"
