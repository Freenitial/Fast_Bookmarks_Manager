<# :
    @echo off & chcp 65001 >nul & cd /d "%~dp0" & Title Edge Favorites
    if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~f0" %* && exit
    powershell /nologo /noprofile /executionpolicy bypass /windowstyle Hidden /command ^
        "&{[ScriptBlock]::Create((gc """%~f0""" -Raw)).Invoke(@(&{$args}%*))}" & exit
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to update the progress bar
function Update-ProgressBar($value) {
    $progressBar.Value = $value
    $form.Refresh()
}

# Confirmation dialog
$result = [System.Windows.Forms.MessageBox]::Show(
    "Edge will be closed and restarted. Do you want to continue?",
    "Confirmation",
    [System.Windows.Forms.MessageBoxButtons]::OKCancel
)

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    # Create and show progress form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Progress"
    $form.Size = New-Object System.Drawing.Size(300, 100)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Size = New-Object System.Drawing.Size(260, 23)
    $progressBar.Location = New-Object System.Drawing.Point(10, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $progressBar.Value = 0

    $form.Controls.Add($progressBar)
    $form.Show()

    # Close Edge
    Update-ProgressBar 10
    taskkill /f /im msedge.exe
    Start-Sleep -Seconds 1

    # Set up paths
    Update-ProgressBar 20
    $bookmarks_file = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft\Edge\User Data\Default\Bookmarks')
    $script_folder = (Get-Location).Path

    # Function to generate a compatible timestamp
    function Get-ChromeTimestamp {
        return [math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date "1601-01-01")).TotalMicroseconds)
    }

    # Check if Bookmarks file exists, create if not
    Update-ProgressBar 30
    if (-not (Test-Path $bookmarks_file)) {
        $initial_bookmarks = @{
            checksum = ""
            roots = @{
                bookmark_bar = @{
                    children = @()
                    date_added = (Get-ChromeTimestamp).ToString()
                    date_last_used = "0"
                    date_modified = (Get-ChromeTimestamp).ToString()
                    guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                    id = "1"
                    name = "Bookmarks Bar"
                    source = "unknown"
                    type = "folder"
                }
                other = @{
                    children = @()
                    date_added = (Get-ChromeTimestamp).ToString()
                    date_last_used = "0"
                    date_modified = "0"
                    guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                    id = "2"
                    name = "Other Bookmarks"
                    source = "unknown"
                    type = "folder"
                }
                synced = @{
                    children = @()
                    date_added = (Get-ChromeTimestamp).ToString()
                    date_last_used = "0"
                    date_modified = "0"
                    guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                    id = "3"
                    name = "Mobile Bookmarks"
                    source = "unknown"
                    type = "folder"
                }
            }
            version = 1
        }
        $initial_bookmarks | ConvertTo-Json -Depth 100 | Set-Content -Path $bookmarks_file -Encoding UTF8
    }

    # Load bookmarks JSON file
    Update-ProgressBar 40
    $bookmarks_data = Get-Content -Path $bookmarks_file -Raw | ConvertFrom-Json

    # Extract all existing URLs in bookmarks
    $existing_urls = @{}
    function Extract-ExistingUrls($node) {
        if ($node.type -eq 'url') {
            $existing_urls[$node.url] = $true
        } elseif ($node.type -eq 'folder' -and $node.children) {
            foreach ($child in $node.children) {
                Extract-ExistingUrls $child
            }
        }
    }
    Extract-ExistingUrls $bookmarks_data.roots.bookmark_bar

    # Recursive function to add bookmarks and folders
    Update-ProgressBar 50
    function Add-BookmarksRecursively($parentNode, $currentPath) {
        Get-ChildItem -Path $currentPath | ForEach-Object {
            if ($_.PSIsContainer) {
                # It's a folder
                $folderName = $_.Name
                $existingFolder = $parentNode.children | Where-Object { $_.name -eq $folderName -and $_.type -eq 'folder' }
                if (-not $existingFolder) {
                    $newFolder = @{
                        date_added = (Get-ChromeTimestamp).ToString()
                        date_last_used = "0"
                        date_modified = (Get-ChromeTimestamp).ToString()
                        guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                        id = (Get-ChromeTimestamp).ToString()
                        name = $folderName
                        source = "user_add"
                        type = "folder"
                        children = @()
                    }
                    $parentNode.children += $newFolder
                    $existingFolder = $newFolder
                }
                Add-BookmarksRecursively $existingFolder $_.FullName
            } elseif ($_.Extension -eq '.url') {
                # It's a .url file
                $content = Get-Content -Path $_.FullName -Encoding UTF8
                $url = ($content | Where-Object { $_ -like 'URL=*' }).Substring(4).Trim()
                
                if (-not $existing_urls.ContainsKey($url)) {
                    $timestamp = Get-ChromeTimestamp
                    $new_favorite = @{
                        date_added = $timestamp.ToString()
                        date_last_used = "0"
                        guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                        id = $timestamp.ToString()
                        name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                        show_icon = $false
                        source = "user_add"
                        type = "url"
                        url = $url
                    }
                    $parentNode.children += $new_favorite
                    $existing_urls[$url] = $true
                }
            }
        }
    }

    # Process bookmarks
    Update-ProgressBar 70
    Add-BookmarksRecursively $bookmarks_data.roots.bookmark_bar $script_folder

    # Save updated bookmarks
    Update-ProgressBar 80
    $bookmarks_data | ConvertTo-Json -Depth 100 | Set-Content -Path $bookmarks_file -Encoding UTF8

    # Restart Edge
    Update-ProgressBar 90
    Start-Process "msedge.exe"
    while (-not (Get-Process msedge | Where-Object { $_.MainWindowTitle -ne "" })) {
        Start-Sleep -Milliseconds 100
    }

    # Close progress form
    $form.Close()
}
