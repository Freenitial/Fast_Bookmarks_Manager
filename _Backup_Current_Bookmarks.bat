<# :
    @echo off & chcp 65001 >nul & cd /d "%~dp0" & Title Browser Bookmarks Export

    set "debug=false"

    if "%debug%"=="true" (set "style=Normal") else (set "style=Hidden")
    powershell /nologo /noprofile /executionpolicy bypass /WindowStyle %style% /command ^
        "&{[ScriptBlock]::Create((gc """%~f0""" -Raw)).Invoke(@(&{$args}%*))}" 
    if "%debug%"=="true" (pause) else (exit)
#>

$OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
$global:SelectedBookmarks = @{}
$global:ExportErrors = @()

# Function to get Chrome profiles
function Get-ChromeProfiles
 {
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
$form.Text = "Browser Bookmarks Export"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 10)

# Create Chrome group box
$chromeGroup = New-Object System.Windows.Forms.GroupBox
$chromeGroup.Location = New-Object System.Drawing.Point(20, 20)
$chromeGroup.Size = New-Object System.Drawing.Size(260, 380)
$chromeGroup.Text = "Chrome"
$form.Controls.Add($chromeGroup)

# Create Edge group box
$edgeGroup = New-Object System.Windows.Forms.GroupBox
$edgeGroup.Location = New-Object System.Drawing.Point(300, 20)
$edgeGroup.Size = New-Object System.Drawing.Size(260, 380)
$edgeGroup.Text = "Edge"
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
            $checkbox.Tag = $web_profile

            $bookmarksFile = if ($browser -eq "Chrome") {
                "$env:LOCALAPPDATA\Google\Chrome\User Data\$web_profile\Bookmarks"
            } else {
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$web_profile\Bookmarks"
            }
            
            if (Test-Path $bookmarksFile) {
                $bookmarkCount = Count-Bookmarks $bookmarksFile
                $selectedCount = 0
                $key = "$browser|$web_profile"
                if ($global:SelectedBookmarks.ContainsKey($key)) {
                    $selectedCount = ($global:SelectedBookmarks[$key] | Where-Object { $_ -ne $null }).Count
                }
                if ($bookmarkCount -eq 0) {
                    $checkbox.Enabled = $false
                    $checkbox.Text += " (No bookmarks)"
                } else {
                    $checkbox.Text += " ($bookmarkCount / $bookmarkCount)"
                    $global:SelectedBookmarks[$key] = Get-AllBookmarks $bookmarksFile
                    
                    # Add View button
                    $viewButton = New-Object System.Windows.Forms.Button
                    $viewButton.Location = New-Object System.Drawing.Point(200, $yPos)
                    $viewButton.Size = New-Object System.Drawing.Size(50, 20)
                    $viewButton.Text = "View"
                    $viewButton.Tag = @{Browser = $browser; Profile = $web_profile; BookmarksFile = $bookmarksFile; Checkbox = $checkbox}
                    $viewButton.Add_Click({
                        $buttonTag = $this.Tag
                        Show-BookmarksTable $buttonTag.Browser $buttonTag.Profile $buttonTag.BookmarksFile $buttonTag.Checkbox
                    })
                    $group.Controls.Add($viewButton)
                }
            } else {
                $checkbox.Enabled = $false
                $checkbox.Text += " (No bookmarks file)"
            }
            
            $group.Controls.Add($checkbox)
            $checkboxes += $checkbox
            $yPos += 30
        }
    }
    return $checkboxes
}

function Get-AllBookmarks($bookmarksFile) {
    $bookmarks = Get-Content -Path $bookmarksFile -Raw -Encoding utf8 | ConvertFrom-Json
    
    function Get-BookmarksRecursive($nodes) {
        $result = @()
        foreach ($node in $nodes) {
            if ($node.type -eq 'url') {
                $result += @{Text = $node.name; Tag = $node.url; Type = "Bookmark"}
            } elseif ($node.children) {
                $childrenResult = Get-BookmarksRecursive $node.children
                if ($childrenResult.Count -gt 0) {
                    $result += @{
                        Text = $node.name
                        Type = "Folder"
                        Children = $childrenResult
                    }
                }
            }
        }
        return $result
    }
    
    return Get-BookmarksRecursive $bookmarks.roots.bookmark_bar.children
}

function Show-BookmarksTable {
    param (
        [string]$browser,
        [string]$profile,
        [string]$bookmarksFile,
        [System.Windows.Forms.CheckBox]$checkbox
    )

    if (-not (Test-Path $bookmarksFile)) {
        [System.Windows.Forms.MessageBox]::Show("Bookmarks file not found: $bookmarksFile", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $bookmarks = Get-Content -Path $bookmarksFile -Raw -Encoding utf8 | ConvertFrom-Json
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Bookmarks - $browser ($profile)"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Arial", 10)

    # Add OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $okButton.Size = New-Object System.Drawing.Size(100, 40)
    $okButton.Location = New-Object System.Drawing.Point(350, 10)
    $okButton.Add_Click({
        $selectedResult = Get-CheckedNodes $treeView.Nodes
        $selectedNodes = $selectedResult.Children
        $selectedCount = $selectedResult.BookmarkCount
        $key = "$browser|$profile"
        $global:SelectedBookmarks[$key] = $selectedNodes
        $totalCount = Count-Bookmarks $bookmarksFile
        $checkbox.Text = "$profile ($selectedCount / $totalCount)"
        $form.Close()
    })
    $form.Controls.Add($okButton)

    # Create TreeView to display bookmarks structure
    $treeView = New-Object System.Windows.Forms.TreeView
    $treeView.Location = New-Object System.Drawing.Point(10, 60)
    $treeView.Size = New-Object System.Drawing.Size(760, 490)
    $treeView.CheckBoxes = $true

    function Add-BookmarksToTreeNode {
        param ($node, $parentTreeNode)

        if ($node.type -eq 'url') {
            $newNode = $parentTreeNode.Nodes.Add([System.Web.HttpUtility]::HtmlDecode($node.name))
            $newNode.Tag = $node.url
            $newNode.Checked = $true
        }

        if ($node.children) {
            $newFolderNode = $parentTreeNode.Nodes.Add($node.name)
            foreach ($child in $node.children) {
                Add-BookmarksToTreeNode $child $newFolderNode
            }
            $newFolderNode.Checked = $true
        }
    }

    $bookmarkBarChildren = $bookmarks.roots.bookmark_bar.children
    foreach ($child in $bookmarkBarChildren) {
        Add-BookmarksToTreeNode $child $treeView
    }
    $treeView.Nodes | ForEach-Object { $_.Checked = $true }
    $treeView.CollapseAll()

    # Attach event handler for checking/unchecking nodes
    $treeView.Add_AfterCheck({
        param ($sender, $e)
        # Propagate check state to children
        foreach ($childNode in $e.Node.Nodes) {
            $childNode.Checked = $e.Node.Checked
        }
    })

    $form.Controls.Add($treeView)
    $form.ShowDialog()
}


function Get-CheckedNodes($nodes) {
    $checkedNodes = @()
    $bookmarkCount = 0

    foreach ($node in $nodes) {
        if ($node.Checked) {
            if ($node.Nodes.Count -eq 0) {
                # C'est un marque-page
                $checkedNodes += @{Text = $node.Text; Tag = $node.Tag; Type = "Bookmark"}
                $bookmarkCount++
            } else {
                # C'est un dossier
                $folderResult = Get-CheckedNodes $node.Nodes
                if ($folderResult.Children.Count -gt 0) {
                    $checkedNodes += @{
                        Text = $node.Text
                        Type = "Folder"
                        Children = $folderResult.Children
                    }
                    $bookmarkCount += $folderResult.BookmarkCount
                }
            }
        }
    }

    return @{
        Children = $checkedNodes
        BookmarkCount = $bookmarkCount
    }
}




$chromeProfiles = Get-ChromeProfiles
$edgeProfiles = Get-EdgeProfiles

$chromeCheckboxes = Create-ProfileCheckboxes $chromeGroup $chromeProfiles "Chrome"
$edgeCheckboxes = Create-ProfileCheckboxes $edgeGroup $edgeProfiles "Edge"

# Create Export button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(250, 420)
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$exportButton.Text = "Export"
$exportButton.Add_Click({
    $selectedProfiles = @()
    if ($chromeProfiles) {
        for ($i = 0; $i -lt $chromeCheckboxes.Count; $i++) {
            if ($chromeCheckboxes[$i].Checked) {
                $selectedProfiles += @{Browser = "Chrome"; Profile = $chromeCheckboxes[$i].Tag}
            }
        }
    }
    if ($edgeProfiles) {
        for ($i = 0; $i -lt $edgeCheckboxes.Count; $i++) {
            if ($edgeCheckboxes[$i].Checked) {
                $selectedProfiles += @{Browser = "Edge"; Profile = $edgeCheckboxes[$i].Tag}
            }
        }
    }
    
    if ($selectedProfiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one profile to export.", "No Profile Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    Write-Host "Selected profiles before export:"
    $selectedProfiles | ForEach-Object { Write-Host "Browser: $($_.Browser), Profile: $($_.Profile)" }
    
    $form.Hide()
    Export-Bookmarks $selectedProfiles
    $form.Close()
})
$form.Controls.Add($exportButton)


function Show-ProgressBar($title, $max) {
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = $title
    $progressForm.Size = New-Object System.Drawing.Size(300, 100)
    $progressForm.StartPosition = "CenterScreen"

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 20)
    $progressBar.Size = New-Object System.Drawing.Size(260, 20)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    # Assurez-vous que Maximum est au moins 1
    $progressBar.Minimum = 0
    $progressBar.Maximum = [Math]::Max(1, $max)
    $progressBar.Value = 0

    $progressForm.Controls.Add($progressBar)

    $progressForm.Show()
    $progressForm.Focus()

    return @{Form = $progressForm; Bar = $progressBar}
}


# Function to export Bookmarks
function Export-Bookmarks($web_profiles) {
    $script_folder = (Get-Location).Path
    Write-Host "Debug: Script folder is $script_folder"
    
    $totalBookmarks = 0
    $web_profiles | ForEach-Object {
        $browser = $_.Browser
        $web_profileName = $_.Profile
        $key = "$browser|$web_profileName"
        Write-Host "Debug: Checking key $key"
        if ($global:SelectedBookmarks.ContainsKey($key)) {
            $count = Count-Bookmarks-Recursive $global:SelectedBookmarks[$key]
            $totalBookmarks += $count
            Write-Host "Debug: Found $count bookmarks for $key"
        } else {
            Write-Host "Debug: No bookmarks found for $key"
        }
    }

    Write-Host "Debug: Total bookmarks to export: $totalBookmarks"

    $totalBookmarks = [Math]::Max(1, $totalBookmarks)
    $progress = Show-ProgressBar "Exporting Bookmarks" $totalBookmarks
    $currentBookmark = 0

    foreach ($web_profile in $web_profiles) {
        $browser = $web_profile.Browser
        $web_profileName = $web_profile.Profile
        $key = "$browser|$web_profileName"

        Write-Host "Debug: Processing $browser profile: $web_profileName"
        if ($global:SelectedBookmarks.ContainsKey($key)) {
            $selectedNodes = $global:SelectedBookmarks[$key]
            Write-Host "Debug: Found $(Count-Bookmarks-Recursive $selectedNodes) selected nodes for $key"
            Process-SelectedBookmarks $selectedNodes $script_folder ([ref]$currentBookmark) $progress
        } else {
            Write-Host "Debug: No bookmarks selected for $browser profile: $web_profileName"
        }
    }

    $progress.Form.Close()

    if ($global:ExportErrors.Count -gt 0) {
        Show-ErrorList
    } else {
        [System.Windows.Forms.MessageBox]::Show("Bookmarks have been exported successfully.", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }

    # Clear the error list for the next export
    $global:ExportErrors = @()
}

function Show-ErrorList {
    $errorForm = New-Object System.Windows.Forms.Form
    $errorForm.Text = "Export Errors"
    $errorForm.Size = New-Object System.Drawing.Size(700, 500)
    $errorForm.StartPosition = "CenterScreen"
    $errorForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(660, 40)
    $label.Text = "The following items could not be exported due to long paths or other errors:"
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $label.ForeColor = [System.Drawing.Color]::Red
    $errorForm.Controls.Add($label)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 50)
    $listView.Size = New-Object System.Drawing.Size(660, 350)
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true

    $listView.Columns.Add("Type", 40) | Out-Null
    $listView.Columns.Add("Path", 490) | Out-Null

    foreach ($error in $global:ExportErrors) {
        if ($error -match "^Failed to create (file|folder): (.+)$") {
            $item = New-Object System.Windows.Forms.ListViewItem($matches[1])
            $item.SubItems.Add($matches[2])
            $listView.Items.Add($item) | Out-Null
        }
    }

    $errorForm.Controls.Add($listView)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(290, 410)
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Text = "OK"
    $okButton.Add_Click({ $errorForm.Close() })
    $errorForm.Controls.Add($okButton)

    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Location = New-Object System.Drawing.Point(10, 410)
    $copyButton.Size = New-Object System.Drawing.Size(100, 30)
    $copyButton.Text = "Copy to Clipboard"
    $copyButton.Add_Click({
        $errorText = $global:ExportErrors -join "`r`n"
        [System.Windows.Forms.Clipboard]::SetText($errorText)
        [System.Windows.Forms.MessageBox]::Show("Errors copied to clipboard.", "Copy Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $errorForm.Controls.Add($copyButton)

    $errorForm.ShowDialog()
}

function Count-Bookmarks-Recursive($nodes) {
    $count = 0
    foreach ($node in $nodes) {
        if ($node.Type -eq "Bookmark") {
            $count++
        } elseif ($node.Type -eq "Folder") {
            $count += Count-Bookmarks-Recursive $node.Children
        }
    }
    return $count
}



function Create-UrlFile($name, $url, $path) {
    if ($url -match '^(chrome|edge)://') {
        Write-Host "Skipping browser-specific URL: $url"
        return
    }

    $content = @"
[InternetShortcut]
URL=$url
"@
    function Get-ValidFileName($fileName, $url) {
        if ($fileName -match '^(https?://)?(www\.)?') { 
            $fileName = $fileName -replace '^(https?://)?(www\.)?', '' 
        }
        $validName = $fileName -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
        $validName = $validName -replace '[\x00-\x1F]', ''
        $validName = $validName.Trim(' ._-')
        if ([string]::IsNullOrWhiteSpace($validName) -or $validName -match '^[\s._-]+$') {
            $validName = $url -replace '^(https?://)?(www\.)?', '' -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
            $validName = $validName.Trim(' ._-')
        }
        if ([string]::IsNullOrWhiteSpace($validName)) {
            $validName = "Unnamed_Bookmark"
        }
        return $validName
    }

    # Check if URL already exists in the current path
    $existingFiles = Get-ChildItem -Path $path -Filter "*.url"
    foreach ($file in $existingFiles) {
        $existingContent = Get-Content -LiteralPath $file.FullName -Raw
        if ($existingContent -match "URL=$([regex]::Escape($url))") {
            Write-Host "Skipping duplicate URL in the same path: $url"
            return
        }
    }

    $validName = Get-ValidFileName $name $url
    $validName = $validName -replace '_+', '_'
    if ($validName.Length -gt 60) {
        $validName = $validName.Substring(0, 60).TrimEnd(' ._-')
    }
    
    $filePath = Join-Path -Path $path -ChildPath "$validName.url"
    $counter = 1

    while (Test-Path -LiteralPath $filePath) {
        $filePath = Join-Path -Path $path -ChildPath "$validName ($counter).url"
        $counter++
    }
    
    try {
        $content | Out-File -LiteralPath $filePath -Encoding utf8 -Force -ErrorAction Stop
        Write-Host "Created URL file: $filePath"
    }
    catch {
        Write-Host "Failed to create URL file: $filePath. Error: $($_.Exception.Message)"
        $global:ExportErrors += "Failed to create file: $filePath"
    }
}

function Process-SelectedBookmarks($nodes, $currentPath, [ref]$currentBookmark, $progress) {
    foreach ($node in $nodes) {
        if ($node.Type -eq "Bookmark") {
            Create-UrlFile $node.Text $node.Tag $currentPath
            $currentBookmark.Value++
            $progress.Bar.Value = $currentBookmark.Value
        } elseif ($node.Type -eq "Folder") {
            $folderName = $node.Text -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
            $folderName = $folderName.Trim(' ._-')
            if ($folderName.Length -gt 60) {
                $folderName = $folderName.Substring(0, 60).TrimEnd(' ._-')
            }
            $folderPath = Join-Path -Path $currentPath -ChildPath $folderName
            
            try {
                if (-not (Test-Path -LiteralPath $folderPath)) {
                    New-Item -Path $folderPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                }
                Process-SelectedBookmarks $node.Children $folderPath $currentBookmark $progress
            }
            catch {
                Write-Host "Failed to create folder: $folderPath. Error: $($_.Exception.Message)"
                $global:ExportErrors += "Failed to create folder: $folderPath"
            }
        }
    }
}

# Show the form
$form.ShowDialog()
