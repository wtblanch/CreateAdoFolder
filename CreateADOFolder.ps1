Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

# === Configuration ===
$config = @{
    Organization = "https://dev.azure.com/YourOrgName"  # Your Azure DevOps organization URL
    Project = "YourProjectName"                         # Your project name
    Repository = "YourRepoName"                         # Your repository name
    SharePointUrls = @(                                # Up to 3 SharePoint URLs
        "https://sharepoint.com/url1",
        "https://sharepoint.com/url2",
        "https://sharepoint.com/url3"
    )
}

# === Create Form ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure DevOps Uploader"
$form.Size = New-Object System.Drawing.Size(600, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize = New-Object System.Drawing.Size(600, 900)
$form.AutoScroll = $true
$form.MaximizeBox = $true

# === Title Label ===
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Azure DevOps File Uploader"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$titleLabel.Size = New-Object System.Drawing.Size(400, 40)
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($titleLabel)

# === Connection Group Box ===
$connectionGroup = New-Object System.Windows.Forms.GroupBox
$connectionGroup.Text = "Connection Settings"
$connectionGroup.Size = New-Object System.Drawing.Size(540, 140)
$connectionGroup.Location = New-Object System.Drawing.Point(20, 70)
$connectionGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($connectionGroup)

# === Project Group Box ===
$projectGroup = New-Object System.Windows.Forms.GroupBox
$projectGroup.Text = "Project Settings"
$projectGroup.Size = New-Object System.Drawing.Size(540, 160)
$projectGroup.Location = New-Object System.Drawing.Point(20, 220)
$projectGroup.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.Controls.Add($projectGroup)

# === Add Controls ===
$labels = @("PAT", "Organization URL", "Project", "Repo", "App Name", "Version")
$textBoxes = @{}

# Split labels into connection and project groups
$connectionLabels = $labels[0..1]  # PAT and Organization URL
$projectLabels = $labels[2..5]     # Project, Repo, App Name, Version

# Add controls to connection group
for ($i = 0; $i -lt $connectionLabels.Length; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $connectionLabels[$i]
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $yPos = 30 + ($i * 45)
    $label.Location = New-Object System.Drawing.Point(20, $yPos)
    $label.Size = New-Object System.Drawing.Size(130, 20)
    $connectionGroup.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(160, $yPos)
    $textbox.Size = New-Object System.Drawing.Size(350, 20)
    $textbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    if ($connectionLabels[$i] -eq "PAT") { 
        $textbox.UseSystemPasswordChar = $true 
    }
    elseif ($connectionLabels[$i] -eq "Organization URL") {
        $textbox.Text = $config.Organization
        $textbox.Enabled = $false  # Make it read-only
    }
    $connectionGroup.Controls.Add($textbox)
    $textBoxes[$connectionLabels[$i]] = $textbox
}

# Add controls to project group
for ($i = 0; $i -lt $projectLabels.Length; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $projectLabels[$i]
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $yPos = 30 + ($i * 30)
    $label.Location = New-Object System.Drawing.Point(20, $yPos)
    $label.Size = New-Object System.Drawing.Size(130, 20)
    $projectGroup.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(160, $yPos)
    $textbox.Size = New-Object System.Drawing.Size(350, 20)
    $textbox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Add tooltips for specific fields
    if ($projectLabels[$i] -eq "Project") {
        $textbox.Text = $config.Project
        $textbox.Enabled = $false  # Make it read-only
    }
    elseif ($projectLabels[$i] -eq "Repo") {
        $textbox.Text = $config.Repository
        $textbox.Enabled = $false  # Make it read-only
    }
    elseif ($projectLabels[$i] -eq "App Name") {
        $textbox.PlaceholderText = "e.g., MyApp"
        $toolTip = New-Object System.Windows.Forms.ToolTip
        $toolTip.SetToolTip($textbox, "Enter your application name (e.g., MyApp, CustomerPortal)")
    }
    elseif ($projectLabels[$i] -eq "Version") {
        $textbox.PlaceholderText = "e.g., POC"
        $toolTip = New-Object System.Windows.Forms.ToolTip
        $toolTip.SetToolTip($textbox, "Enter version or stage (e.g., POC, MVP, Version 2)")
    }

    $projectGroup.Controls.Add($textbox)
    $textBoxes[$projectLabels[$i]] = $textbox
}

# === SharePoint URL Input ===
$spGroup = New-Object System.Windows.Forms.GroupBox
$spGroup.Text = "SharePoint URLs"
$spGroup.Size = New-Object System.Drawing.Size(540, 140)
$spGroup.Location = New-Object System.Drawing.Point(20, 390)
$form.Controls.Add($spGroup)

$spLabel = New-Object System.Windows.Forms.Label
$spLabel.Text = "Enter SharePoint URLs (one per line):"
$spLabel.Location = New-Object System.Drawing.Point(20, 30)
$spLabel.Size = New-Object System.Drawing.Size(200, 20)
$spGroup.Controls.Add($spLabel)

$spTextBox = New-Object System.Windows.Forms.TextBox
$spTextBox.Multiline = $true
$spTextBox.ScrollBars = "Vertical"
$spTextBox.Size = New-Object System.Drawing.Size(420, 70)
$spTextBox.Location = New-Object System.Drawing.Point(20, 50)
$spTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Pre-fill SharePoint URLs from config
if ($config.SharePointUrls) {
    $spTextBox.Text = $config.SharePointUrls -join "`r`n"
}

# Add URL validation button
$validateButton = New-Object System.Windows.Forms.Button
$validateButton.Text = "Validate URLs"
$validateButton.Size = New-Object System.Drawing.Size(80, 70)
$validateButton.Location = New-Object System.Drawing.Point(440, 50)
$validateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$validateButton.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$validateButton.Add_Click({
    $urls = $spTextBox.Text -split "`r`n" | Where-Object { $_ -ne "" }
    foreach ($url in $urls) {
        try {
            if ($url -match '^https:\/\/.*\.sharepoint\.com.*$') {
                $statusBox.AppendText("Valid SharePoint URL: $url`r`n")
            } else {
                $statusBox.AppendText("Invalid SharePoint URL format: $url`r`n")
            }
        } catch {
            $statusBox.AppendText("Error validating URL: $url`r`n")
        }
    }
})
$spGroup.Controls.Add($validateButton)
$spGroup.Controls.Add($spTextBox)

# === File Selection Group ===
$fileGroup = New-Object System.Windows.Forms.GroupBox
$fileGroup.Text = "Local Files and Folders"
$fileGroup.Size = New-Object System.Drawing.Size(540, 200)
$fileGroup.Location = New-Object System.Drawing.Point(20, 540)
$form.Controls.Add($fileGroup)

# Create ListView for selected items
$listView = New-Object System.Windows.Forms.ListView
$listView.View = [System.Windows.Forms.View]::Details
$listView.Size = New-Object System.Drawing.Size(500, 120)
$listView.Location = New-Object System.Drawing.Point(20, 60)
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Type", 60) | Out-Null
$listView.Columns.Add("Name", 250) | Out-Null
$listView.Columns.Add("Path", 190) | Out-Null
$fileGroup.Controls.Add($listView)

# File and folder arrays
$localFiles = @()
$localFolders = @()

# Add Files Button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Add Files"
$browseButton.Location = New-Object System.Drawing.Point(20, 25)
$browseButton.Size = New-Object System.Drawing.Size(245, 30)
$browseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$browseButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$browseButton.ForeColor = [System.Drawing.Color]::White
$browseButton.Add_Click({
    $openFile = New-Object System.Windows.Forms.OpenFileDialog
    $openFile.Multiselect = $true
    $openFile.Filter = "All Files (*.*)|*.*"
    if ($openFile.ShowDialog() -eq "OK") {
        foreach ($file in $openFile.FileNames) {
            if ($file -notin $localFiles) {
                $localFiles += $file
                $item = New-Object System.Windows.Forms.ListViewItem("File")
                $item.SubItems.Add([System.IO.Path]::GetFileName($file))
                $item.SubItems.Add($file)
                $listView.Items.Add($item)
            }
        }
        $statusBox.AppendText("Added $($openFile.FileNames.Count) files`r`n")
    }
})
$fileGroup.Controls.Add($browseButton)

# Add Folder Button
$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Text = "Add Folder"
$folderButton.Location = New-Object System.Drawing.Point(275, 25)
$folderButton.Size = New-Object System.Drawing.Size(245, 30)
$folderButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$folderButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$folderButton.ForeColor = [System.Drawing.Color]::White
$folderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder to add"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        if ($folderBrowser.SelectedPath -notin $localFolders) {
            $localFolders += $folderBrowser.SelectedPath
            $item = New-Object System.Windows.Forms.ListViewItem("Folder")
            $item.SubItems.Add([System.IO.Path]::GetFileName($folderBrowser.SelectedPath))
            $item.SubItems.Add($folderBrowser.SelectedPath)
            $listView.Items.Add($item)
        }
        $statusBox.AppendText("Added folder: $($folderBrowser.SelectedPath)`r`n")
    }
})
$fileGroup.Controls.Add($folderButton)

# Remove Button
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "Remove Selected"
$removeButton.Location = New-Object System.Drawing.Point(20, 160)
$removeButton.Size = New-Object System.Drawing.Size(500, 30)
$removeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$removeButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$removeButton.ForeColor = [System.Drawing.Color]::White
$removeButton.Add_Click({
    $selected = $listView.SelectedItems
    foreach ($item in $selected) {
        $path = $item.SubItems[2].Text
        if ($item.SubItems[0].Text -eq "File") {
            $localFiles = $localFiles | Where-Object { $_ -ne $path }
        } else {
            $localFolders = $localFolders | Where-Object { $_ -ne $path }
        }
        $listView.Items.Remove($item)
    }
    $statusBox.AppendText("Removed $($selected.Count) items`r`n")
})
$fileGroup.Controls.Add($removeButton)

# === Upload to ADO Button Logic ===

$uploadButton = New-Object System.Windows.Forms.Button
$uploadButton.Text = "Upload to ADO"
$uploadButton.Size = New-Object System.Drawing.Size(540, 40)
$uploadButton.Location = New-Object System.Drawing.Point(20, 750)
$uploadButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$uploadButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$uploadButton.ForeColor = [System.Drawing.Color]::White

$uploadButton.Add_Click({
    try {
        $statusBox.AppendText("Starting upload process...`r`n")

        $organizationUrl = $textBoxes['Organization URL'].Text.TrimEnd('/')
        $project = $textBoxes['Project'].Text        
        $repo = $textBoxes['Repo'].Text
        Add-Type -AssemblyName System.Web
        $encodedRepo = [System.Web.HttpUtility]::UrlEncode($repo)    
        $pat = $textBoxes['PAT'].Text
        $appName = $textBoxes['App Name'].Text
        $version = $textBoxes['Version'].Text

        if (-not $pat) {
            [System.Windows.Forms.MessageBox]::Show("PAT token is required.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
        $headers = @{ Authorization = "Basic $base64AuthInfo" }

        $branchName = "refs/heads/main"  # Change if needed

        # Step 1: Get latest commit
        $refsUrl = "$organizationUrl/$project/_apis/git/repositories/$encodedRepo/refs?filter=heads/main&api-version=6.0"
        $refResponse = Invoke-RestMethod -Uri $refsUrl -Headers $headers -Method Get -ErrorAction Stop
        if (-not $refResponse.value) {
            throw "Branch not found or no refs returned."
        }
        $latestCommit = $refResponse.value[0].objectId
        $statusBox.AppendText("Fetched latest commit.`r`n")

        # Step 2: Prepare changes
        $changes = @()
        foreach ($item in $listView.Items) {
            if ($item.Text -eq "File") {
                $filePath = $item.SubItems[2].Text
                if (!(Test-Path $filePath)) {
                    $statusBox.AppendText("Warning: File not found - $filePath`r`n")
                    continue
                }
                $contentBytes = [System.IO.File]::ReadAllBytes($filePath)
                $contentString = [System.Text.Encoding]::UTF8.GetString($contentBytes)
                $relativePath = "$appName/$version/$(Split-Path $filePath -Leaf)"

                $changes += @{
                    changeType = "add"
                    item = @{ path = "/$relativePath" }
                    newContent = @{
                        content = $contentString
                        contentType = "rawtext"
                    }
                }
            }
        }

        if ($changes.Count -eq 0) {
            throw "No valid files to upload."
        }

        $statusBox.AppendText("Prepared $($changes.Count) file changes.`r`n")

        # Step 3: Push changes
        $pushBody = @{
            refUpdates = @(@{
                name = $branchName
                oldObjectId = $latestCommit
            })
            commits = @(@{
                comment = "Creating new folder for $appName"
                changes = $changes
            })
        } | ConvertTo-Json -Depth 10
        $pushUrl = "$organizationUrl/$project/_apis/git/repositories/$encodedRepo/pushes?&api-version=6.0"
        $statusBox.AppendText("Pushing changes to: $pushUrl`r`n")
        $pushResponse = Invoke-RestMethod -Uri $pushUrl -Headers $headers -Method Post -Body $pushBody -ContentType "application/json" -ErrorAction Stop
        $statusBox.AppendText("Upload completed successfully!`r`n")
    }
    catch {
        if ($_.Exception -and $_.Exception.Message) {
            $errorMessage = $_.Exception.Message
            $statusBox.AppendText("Error: $errorMessage`r`n")
            
            # Try to pull the URL that failed if possible
            if ($_.Exception.InnerException -and $_.Exception.InnerException.Message) {
                $statusBox.AppendText("Inner Exception: $($_.Exception.InnerException.Message)`r`n")
            }
        } else {
            $statusBox.AppendText("Error: $_`r`n")
        }
    
        [System.Windows.Forms.MessageBox]::Show("Upload failed. See status window.", "Upload Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }   
})

$form.Controls.Add($uploadButton)


# === Status Box ===
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Size = New-Object System.Drawing.Size(540, 60)
$statusBox.Location = New-Object System.Drawing.Point(20, 800)
$statusBox.ReadOnly = $true
$statusBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($statusBox)

# === Show Form ===
[void]$form.ShowDialog()
