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
                if ($bookmarkCount -eq 0) {
                    $checkbox.Enabled = $false
                    $checkbox.Text += " (No bookmarks)"
                } else {
                    $checkbox.Text += " ($bookmarkCount)"
                    
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
    
    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("Name") | Out-Null
    $dataTable.Columns.Add("URL") | Out-Null

    function Add-BookmarksToTable($node) {
        if ($node.type -eq 'url') {
            $row = $dataTable.NewRow()
            $row["Name"] = [System.Web.HttpUtility]::HtmlDecode($node.name)
            $row["URL"] = $node.url
            $dataTable.Rows.Add($row)
        }
        if ($node.children) {
            foreach ($child in $node.children) {
                Add-BookmarksToTable $child
            }
        }
    }

    Add-BookmarksToTable $bookmarks.roots.bookmark_bar

    # Create and show the DataGridView form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Bookmarks - $browser ($profile)"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Arial", 10)

    $dataGridView = New-Object System.Windows.Forms.DataGridView
    $dataGridView.Dock = "Fill"
    $dataGridView.DataSource = $dataTable
    $dataGridView.AutoSizeColumnsMode = "Fill"
    $dataGridView.RowHeadersVisible = $false
    $dataGridView.AllowUserToAddRows = $false
    $dataGridView.ReadOnly = $true
    $dataGridView.ColumnHeadersHeightSizeMode = 'DisableResizing'
    $dataGridView.RowHeadersWidthSizeMode = 'DisableResizing'
    $dataGridView.AllowUserToResizeRows = $false
    $dataGridView.MultiSelect = $false
    $dataGridView.SelectionMode = 'CellSelect'

    $form.Add_Shown({
        $dataGridView.ClearSelection()
    })

    $dataGridView.Add_CellClick({
        param ($sender, $e)
        $sender.ClearSelection()
    })

    $form.Controls.Add($dataGridView)
    $form.ShowDialog()
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
    $progressBar.Minimum = 0
    $progressBar.Maximum = $max
    $progressBar.Value = 0
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    $progressForm.Controls.Add($progressBar)

    $progressForm.Show()
    $progressForm.Focus()

    return @{Form = $progressForm; Bar = $progressBar}
}


# Function to export Bookmarks
function Export-Bookmarks($web_profiles) {
    $script_folder = (Get-Location).Path
    
    $totalBookmarks = 0
    $web_profiles | ForEach-Object {
        $browser = $_.Browser
        $web_profileName = $_.Profile
        $bookmarks_file = if ($browser -eq "Chrome") {
            "$env:LOCALAPPDATA\Google\Chrome\User Data\$web_profileName\Bookmarks"
        } else {
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$web_profileName\Bookmarks"
        }
        if (Test-Path $bookmarks_file) {
            $bookmarks_data = Get-Content -Path $bookmarks_file -Raw | ConvertFrom-Json
            $totalBookmarks += (Count-Bookmarks-Recursive $bookmarks_data.roots.bookmark_bar)
        }
    }

    $progress = Show-ProgressBar "Exporting Bookmarks" $totalBookmarks
    $currentBookmark = 0

    foreach ($web_profile in $web_profiles) {
        $browser = $web_profile.Browser
        $web_profileName = $web_profile.Profile

        $bookmarks_file = if ($browser -eq "Chrome") {
            "$env:LOCALAPPDATA\Google\Chrome\User Data\$web_profileName\Bookmarks"
        } else {
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$web_profileName\Bookmarks"
        }
        
        if (Test-Path $bookmarks_file) {
            $bookmarks_data = Get-Content -Path $bookmarks_file -Raw | ConvertFrom-Json
            Process-Bookmarks $bookmarks_data.roots.bookmark_bar $script_folder ([ref]$currentBookmark) $progress
        } else {
            Write-Host "Bookmarks file not found: $bookmarks_file"
        }
    }

    $progress.Form.Close()
    [System.Windows.Forms.MessageBox]::Show("Bookmarks have been exported successfully.", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}


function Count-Bookmarks-Recursive($node) {
    $count = 0
    if ($node.type -eq 'url') {
        $count++
    }
    if ($node.children) {
        foreach ($child in $node.children) {
            $count += Count-Bookmarks-Recursive $child
        }
    }
    return $count
}



# Function to create .url file
function Create-UrlFile($name, $url, $path) {
    # Ignorer les URLs qui commencent par chrome:// ou edge://
    if ($url -match '^(chrome|edge)://') {
        Write-Host "Skipping browser-specific URL: $url"
        return
    }

    $content = @"
[InternetShortcut]
URL=$url
"@
    
    # Fonction pour créer un nom de fichier valide
    function Get-ValidFileName($fileName, $url) {
        # Si le nom ressemble à une URL, retirer les préfixes
        if ($fileName -match '^(https?://)?(www\.)?') {
            $fileName = $fileName -replace '^(https?://)?(www\.)?', ''
        }

        # Remplacer les caractères non-alphanumériques (sauf les espaces, @, et apostrophes) par des underscores
        $validName = $fileName -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
        
        # Supprimer les caractères de contrôle
        $validName = $validName -replace '[\x00-\x1F]', ''
        
        # Supprimer les espaces, points, underscores et tirets au début et à la fin
        $validName = $validName.Trim(' ._-')
        
        # Si le nom est vide ou ne contient que des caractères spéciaux, utiliser l'URL
        if ([string]::IsNullOrWhiteSpace($validName) -or $validName -match '^[\s._-]+$') {
            $validName = $url -replace '^(https?://)?(www\.)?', '' -replace '[^\p{L}\p{Nd}\s@''._-]', '_'
            $validName = $validName.Trim(' ._-')
        }
        
        # S'assurer que le nom n'est toujours pas vide
        if ([string]::IsNullOrWhiteSpace($validName)) {
            $validName = "Unnamed_Bookmark"
        }
        
        return $validName
    }
    
    $validName = Get-ValidFileName $name $url
    
    # Remplacer les underscores multiples par un seul underscore
    $validName = $validName -replace '_+', '_'
    
    # Limiter la longueur du nom à 40 caractères
    if ($validName.Length -gt 40) {
        $validName = $validName.Substring(0, 40).TrimEnd(' ._-')
    }
    
    $filePath = Join-Path -Path $path -ChildPath "$validName.url"
    
    try {
        $content | Out-File -FilePath $filePath -Encoding utf8 -Force -ErrorAction Stop
        Write-Host "Created/Overwritten URL file: $filePath"
    }
    catch {
        Write-Host "Failed to create/overwrite URL file: $filePath. Error: $($_.Exception.Message)"
    }
}


# Recursive function to process bookmarks and create folders/files
function Process-Bookmarks($node, $currentPath, [ref]$currentBookmark, $progress) {
    if ($node.children) {
        foreach ($child in $node.children) {
            if ($child.type -eq 'folder') {
                $folderPath = Join-Path -Path $currentPath -ChildPath $child.name
                if (-not (Test-Path $folderPath)) {
                    New-Item -Path $folderPath -ItemType Directory | Out-Null
                }
                Process-Bookmarks $child $folderPath $currentBookmark $progress
            } elseif ($child.type -eq 'url') {
                Create-UrlFile $child.name $child.url $currentPath
                $currentBookmark.Value++
                $progress.Bar.Value = $currentBookmark.Value
                $progress.Form.Refresh()
            }
        }
    } elseif ($node.type -eq 'url') {
        Create-UrlFile $node.name $node.url $currentPath
        $currentBookmark.Value++
        $progress.Bar.Value = $currentBookmark.Value
        $progress.Form.Refresh()
    }
}

# Show the form
$form.ShowDialog()
