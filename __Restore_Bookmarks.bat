<# :
    @echo off & chcp 65001 >nul & cd /d "%~dp0" & Title Edge Favorites
    powershell /nologo /noprofile /executionpolicy bypass /windowstyle Hidden /command ^
        "&{[ScriptBlock]::Create((gc """%~f0""" -Raw)).Invoke(@(&{$args}%*))}" & exit
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get Chrome profiles
function Get-ChromeProfiles {
    $chromeUserData = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    if (-not (Test-Path $chromeUserData)) { return $null }
    $web_profiles = @("Default")
    $web_profiles += Get-ChildItem -Path $chromeUserData -Directory | Where-Object { $_.Name -match "^Profile \d+$" } | ForEach-Object { $_.Name }
    return $web_profiles
}

# Function to get Edge profiles
function Get-EdgeProfiles {
    $edgeUserData = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
    if (-not (Test-Path $edgeUserData)) { return $null }
    $web_profiles = @("Default")
    $web_profiles += Get-ChildItem -Path $edgeUserData -Directory | Where-Object { $_.Name -match "^Profile \d+$" } | ForEach-Object { $_.Name }
    return $web_profiles
}

# Function to count bookmarks
function Count-Bookmarks($bookmarksFile) {
    if (-not (Test-Path $bookmarksFile)) { return 0 }
    $bookmarks = Get-Content -Path $bookmarksFile -Raw -Encoding utf8 | ConvertFrom-Json
    
    function Count-Recursive($node) {
        $count = 0
        if ($node.type -eq 'url') { 
            $count++ 
        }
        if ($node.children) { 
            foreach ($child in $node.children) {
                $count += Count-Recursive $child
            }
        }
        return $count
    }
    return Count-Recursive $bookmarks.roots.bookmark_bar
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Browser Bookmarks Import"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 10)

# Create Chrome group box
$chromeGroup = New-Object System.Windows.Forms.GroupBox
$chromeGroup.Location = New-Object System.Drawing.Point(20, 20)
$chromeGroup.Size = New-Object System.Drawing.Size(260, 380)
$chromeGroup.Text = "Chrome Profiles"
$form.Controls.Add($chromeGroup)

# Create Edge group box
$edgeGroup = New-Object System.Windows.Forms.GroupBox
$edgeGroup.Location = New-Object System.Drawing.Point(300, 20)
$edgeGroup.Size = New-Object System.Drawing.Size(260, 380)
$edgeGroup.Text = "Edge Profiles"
$form.Controls.Add($edgeGroup)

function Create-ProfileCheckboxes($group, $web_profiles, $browser) {
    $yPos = 30
    $checkboxes = @()
    if ($web_profiles -eq $null) {
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $label.Size = New-Object System.Drawing.Size(240, 20)
        $label.Text = "$browser not found"
        $group.Controls.Add($label)
    } else {
        foreach ($web_profile in $web_profiles) {
            $checkbox = New-Object System.Windows.Forms.CheckBox
            $checkbox.Location = New-Object System.Drawing.Point(10, $yPos)
            $checkbox.Size = New-Object System.Drawing.Size(180, 20)
            $checkbox.Text = $web_profile
            $checkbox.Tag = @{Browser = $browser; Profile = $web_profile}

            $bookmarksFile = if ($browser -eq "Chrome") {
                "$env:LOCALAPPDATA\Google\Chrome\User Data\$web_profile\Bookmarks"
            } else {
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$web_profile\Bookmarks"
            }
            
            if (Test-Path $bookmarksFile) {
                $bookmarkCount = Count-Bookmarks $bookmarksFile
                if ($bookmarkCount -eq 0) {
                    $checkbox.Text += " (No bookmarks)"
                } else {
                    $checkbox.Text += " ($bookmarkCount)"
                }
                
                # Add View button
                $viewButton = New-Object System.Windows.Forms.Button
                $viewButton.Location = New-Object System.Drawing.Point(200, $yPos)
                $viewButton.Size = New-Object System.Drawing.Size(50, 20)
                $viewButton.Text = "View"
                $viewButton.Tag = @{Browser = $browser; Profile = $web_profile; BookmarksFile = $bookmarksFile}
                $viewButton.Add_Click({
                    $buttonTag = $this.Tag
                    Show-BookmarksTable $buttonTag.Browser $buttonTag.Profile $buttonTag.BookmarksFile
                })
                $group.Controls.Add($viewButton)
            } else {
                $checkbox.Text += " (No bookmarks file)"
            }
            
            $group.Controls.Add($checkbox)
            $checkboxes += $checkbox
            $yPos += 30
        }
    }
    return $checkboxes
}

function Show-BookmarksTable {
    param (
        [string]$browser,
        [string]$profile,
        [string]$bookmarksFile
    )

    if (-not (Test-Path $bookmarksFile)) {
        [System.Windows.Forms.MessageBox]::Show("Bookmarks file not found: $bookmarksFile", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $bookmarks = Get-Content -Path $bookmarksFile -Raw -Encoding utf8 | ConvertFrom-Json
    
    # Create main form
    $viewForm = New-Object System.Windows.Forms.Form
    $viewForm.Text = "Bookmarks - $browser ($profile)"
    $viewForm.Size = New-Object System.Drawing.Size(800, 600)
    $viewForm.StartPosition = "CenterScreen"
    $viewForm.Font = New-Object System.Drawing.Font("Arial", 10)

    # Create TreeView to display bookmarks structure
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(10, 10)
    $treeView.Size = New-Object System.Drawing.Size(760, 540)

    function Add-BookmarksToTreeNode {
        param ($node, $parentTreeNode)

        if ($node.type -eq 'url') {
            $newNode = $parentTreeNode.Nodes.Add($node.name)
            $newNode.Tag = $node.url
        }

        if ($node.children) {
            $newFolderNode = $parentTreeNode.Nodes.Add($node.name)
            foreach ($child in $node.children) {
                Add-BookmarksToTreeNode $child $newFolderNode
            }
        }
    }

    $bookmarkBarChildren = $bookmarks.roots.bookmark_bar.children
    foreach ($child in $bookmarkBarChildren) {
        Add-BookmarksToTreeNode $child $treeView
    }

    $viewForm.Controls.Add($treeView)
    $viewForm.ShowDialog()
}

# Create checkboxes for Chrome and Edge profiles
$chromeProfiles = Get-ChromeProfiles
$edgeProfiles = Get-EdgeProfiles

$chromeCheckboxes = Create-ProfileCheckboxes $chromeGroup $chromeProfiles "Chrome"
$edgeCheckboxes = Create-ProfileCheckboxes $edgeGroup $edgeProfiles "Edge"

# Add Import button
$importButton = New-Object System.Windows.Forms.Button
$importButton.Location = New-Object System.Drawing.Point(250, 420)
$importButton.Size = New-Object System.Drawing.Size(100, 30)
$importButton.Text = "Import"
$importButton.Add_Click({
    $selectedProfiles = @()
    $chromeCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $selectedProfiles += $_.Tag }
    $edgeCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $selectedProfiles += $_.Tag }

    if ($selectedProfiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile to import to.", "No Profile Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Selected browsers will be closed and restarted. Do you want to continue?",
        "Confirmation",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Create and show progress form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Import Progress"
        $progressForm.Size = New-Object System.Drawing.Size(300, 100)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.TopMost = $true

        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(260, 23)
        $progressBar.Location = New-Object System.Drawing.Point(10, 20)
        $progressBar.Minimum = 0
        $progressBar.Maximum = 100
        $progressBar.Value = 0

        $progressForm.Controls.Add($progressBar)
        $progressForm.Show()

        # Function to update the progress bar
        function Update-ProgressBar($value) {
            $progressBar.Value = $value
            $progressForm.Refresh()
        }

        # Close selected browsers
        Update-ProgressBar 10
        if ($selectedProfiles | Where-Object { $_.Browser -eq "Chrome" }) {
            taskkill /f /im chrome.exe
        }
        if ($selectedProfiles | Where-Object { $_.Browser -eq "Edge" }) {
            taskkill /f /im msedge.exe
        }
        Start-Sleep -Seconds 1

        # Import bookmarks
        $script_folder = (Get-Location).Path
        foreach ($profile in $selectedProfiles) {
            $bookmarks_file = if ($profile.Browser -eq "Chrome") {
                "$env:LOCALAPPDATA\Google\Chrome\User Data\$($profile.Profile)\Bookmarks"
            } else {
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$($profile.Profile)\Bookmarks"
            }

            Update-ProgressBar 30

            # Ensure Bookmarks file exists
            Ensure-BookmarksFile $bookmarks_file

            # Load existing bookmarks
            $bookmarks_data = Get-Content -Path $bookmarks_file -Raw -Encoding UTF8 | ConvertFrom-Json

            # Add new bookmarks
            Update-ProgressBar 50
            Add-BookmarksRecursively $bookmarks_data.roots.bookmark_bar $script_folder

            # Save updated bookmarks
            Update-ProgressBar 70
            $compressedJson = $bookmarks_data | ConvertTo-Json -Depth 100 -Compress
            [System.IO.File]::WriteAllText($bookmarks_file, $compressedJson, [System.Text.Encoding]::UTF8)
        }

        # Restart browsers
        Update-ProgressBar 90
        if ($selectedProfiles | Where-Object { $_.Browser -eq "Chrome" }) {
            Start-Process "chrome.exe"
        }
        if ($selectedProfiles | Where-Object { $_.Browser -eq "Edge" }) {
            Start-Process "msedge.exe"
        }

        # Close progress form
        $progressForm.Close()

        [System.Windows.Forms.MessageBox]::Show("Import completed successfully!", "Import Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($importButton)

# Functions for importing bookmarks
function Get-ChromeTimestamp {
    return [math]::Round((Get-Date).ToUniversalTime().Subtract((Get-Date "1601-01-01")).TotalMicroseconds)
}

function Get-ShortPath($longPath) {
    $obj = New-Object -ComObject Shell.Application
    $folder = $obj.Namespace([System.IO.Path]::GetDirectoryName($longPath))
    $file = $folder.ParseName([System.IO.Path]::GetFileName($longPath))
    return $file.Path
}


function Ensure-BookmarksFile($bookmarksFile) {
    if (-not (Test-Path $bookmarksFile)) {
        $initial_bookmarks = @{
            roots = @{
                bookmark_bar = @{ children = @(); type = "folder" }
                other = @{ children = @(); type = "folder" }
                synced = @{ children = @(); type = "folder" }
            }
            version = 1
        }
        $initial_bookmarks | ConvertTo-Json -Depth 100 | Set-Content -Path $bookmarksFile -Encoding UTF8
    }
}


function Add-BookmarksRecursively($parentNode, $currentPath) {
    Get-ChildItem -Path $currentPath -ErrorAction SilentlyContinue | ForEach-Object {
        $shortPath = Get-ShortPath $_.FullName
        if ($_.PSIsContainer) {
            # It's a folder
            $folderName = $_.Name
            $existingFolder = $parentNode.children | Where-Object { $_.name -eq $folderName -and $_.type -eq 'folder' }
            if (-not $existingFolder) {
                $newFolder = [PSCustomObject]@{
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
            Add-BookmarksRecursively $existingFolder $shortPath
        } elseif ($_.Extension -eq '.url') {
            # It's a .url file
            $content = Get-Content -Path $shortPath -Encoding UTF8 -ErrorAction SilentlyContinue
            if ($content) {
                $url = ($content | Where-Object { $_ -like 'URL=*' }).Substring(4).Trim()
                $name = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                
                # Check if URL already exists in the current folder
                $existingBookmark = $parentNode.children | Where-Object { $_.url -eq $url -and $_.type -eq 'url' }
                if (-not $existingBookmark) {
                    $timestamp = Get-ChromeTimestamp
                    $new_favorite = [PSCustomObject]@{
                        date_added = $timestamp.ToString()
                        date_last_used = "0"
                        guid = "new-guid-{0}" -f [guid]::NewGuid().ToString().Substring(0, 8)
                        id = $timestamp.ToString()
                        name = $name
                        show_icon = $false
                        source = "user_add"
                        type = "url"
                        url = $url
                    }
                    $parentNode.children += $new_favorite
                }
            }
        }
    }
}

$form.ShowDialog()
