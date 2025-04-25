Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module PnP.PowerShell -ErrorAction SilentlyContinue

# === Configuration ===
$config = @{
    Organization = "https://dev.azure.com/YourOrgName"  # Replace with your organization URL
    Project = "YourProjectName"                         # Replace with your project name
    Repository = "YourRepoName"                         # Replace with your repository name
    SharePointUrls = @(                                 # Add up to 3 SharePoint URLs
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

# === Status Box ===
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Size = New-Object System.Drawing.Size(540, 60)
$statusBox.Location = New-Object System.Drawing.Point(20, 750)
$statusBox.ReadOnly = $true
$statusBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$statusBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($statusBox)

# === Start Button ===
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start Upload"
$startButton.Location = New-Object System.Drawing.Point(20, 820)
$startButton.Size = New-Object System.Drawing.Size(540, 35)
$startButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$startButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

$startButton.Add_Click({
    $statusBox.Clear()

    # === Validate Required Fields ===
    foreach ($label in $labels) {
        if ([string]::IsNullOrWhiteSpace($textBoxes[$label].Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter: $label", "Missing Field", "OK", "Error")
            return
        }
    }

    $pat      = $textBoxes["PAT"].Text
    $orgUrl   = $textBoxes["Organization URL"].Text.TrimEnd("/")
    $project  = $textBoxes["Project"].Text
    $repo     = $textBoxes["Repo"].Text
    $appName  = $textBoxes["App Name"].Text
    $version  = $textBoxes["Version"].Text

    # === Create API headers ===
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
    $headers = @{
        Authorization = "Basic $base64AuthInfo"
        "Content-Type" = "application/json"
    }

    try {
        # Create temporary directory
        $tempDir = Join-Path $env:TEMP "$appName-$version"
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Download SharePoint files
        $urls = $spTextBox.Text -split "`r`n" | Where-Object { $_ -ne "" }
        foreach ($url in $urls) {
            $statusBox.AppendText("Processing SharePoint URL: $url`r`n")
            Get-SharePointFiles -Url $url -DestinationPath $tempDir
        }
        
        # Copy local files
        foreach ($file in $localFiles) {
            $fileName = Split-Path $file -Leaf
            Copy-Item -Path $file -Destination (Join-Path $tempDir $fileName)
            $statusBox.AppendText("Copied local file: $fileName`r`n")
        }

        # Generate README.md
        $readmeContent = @"
# $appName

## Version: $version

This repository contains the files and configuration for $appName.

## Repository Information
- **Organization:** $orgUrl
- **Project:** $project
- **Repository:** $repo

## Repository Structure
\`\`\`
$appName/
└── $version/
    ├── README.md
    ├── [SharePoint Files]
$((Get-ChildItem -Path $tempDir -Recurse -Directory | ForEach-Object {
    $indent = "    │   " + ("│   " * ($_.FullName.Split([IO.Path]::DirectorySeparatorChar).Count - $tempDir.Split([IO.Path]::DirectorySeparatorChar).Count))
    "$indent└── $($_.Name)/"
}) -join "`r`n")
$((Get-ChildItem -Path $tempDir -Recurse -File | ForEach-Object {
    $indent = "    │   " + ("│   " * ($_.Directory.FullName.Split([IO.Path]::DirectorySeparatorChar).Count - $tempDir.Split([IO.Path]::DirectorySeparatorChar).Count))
    "$indent└── $($_.Name)"
}) -join "`r`n")
\`\`\`

## Contents
This folder contains:
- Configuration files
- SharePoint documents
- Application-specific settings

## SharePoint Documents
The following SharePoint sites are included:
$(($urls | ForEach-Object { "- $_" }) -join "`r`n")

## Files
$((Get-ChildItem -Path $tempDir -Recurse -File | ForEach-Object { 
    $relativePath = $_.FullName.Replace($tempDir, '').TrimStart('\')
    "- $relativePath"
}) -join "`r`n")

## Version History
- Current Version: $version
  - Initial setup and configuration
  - SharePoint documents synchronized
  - Local files added

## Contact
For questions or issues, please contact the repository maintainers.

---
Last Updated: $(Get-Date -Format "yyyy-MM-dd")
"@

        Set-Content -Path (Join-Path $tempDir "README.md") -Value $readmeContent
        $statusBox.AppendText("Created README.md`r`n")

        # Batch commit all files
        $commitMessage = "Initial commit for $appName $version - Added README.md and synchronized files from SharePoint"
        $success = Submit-BatchCommit -OrgUrl $orgUrl -Project $project -Repo $repo -Headers $headers `
                                    -AppName $appName -Version $version -SourcePath $tempDir `
                                    -CommitMessage $commitMessage

        if ($success) {
            $statusBox.AppendText("Upload completed successfully.`r`n")
        }
        
        # Cleanup
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
            $statusBox.AppendText("Cleaned up temporary directory`r`n")
        }
    }
    catch {
        $statusBox.AppendText("Error during upload: $_`r`n")
    }
})

function Get-SharePointFiles {
    param (
        $Url,
        $DestinationPath
    )
    
    try {
        # Note: This requires manual authentication through browser
        Connect-PnPOnline -Url $Url -Interactive
        
        # Get all files from the SharePoint site
        $files = Get-PnPListItem -List "Documents" -Fields "FileLeafRef", "FileRef"
        
        foreach ($file in $files) {
            $fileName = $file["FileLeafRef"]
            $filePath = $file["FileRef"]
            
            # Create the destination directory if it doesn't exist
            $destDir = Join-Path $DestinationPath (Split-Path $filePath -Parent)
            if (!(Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            
            # Download the file
            Get-PnPFile -Url $filePath -Path $destDir -Filename $fileName -AsFile
            $statusBox.AppendText("Downloaded: $fileName`r`n")
        }
    }
    catch {
        $statusBox.AppendText("Error downloading from SharePoint: $_`r`n")
    }
}

function Submit-BatchCommit {
    param (
        $OrgUrl,
        $Project,
        $Repo,
        $Headers,
        $AppName,
        $Version,
        $SourcePath,
        $CommitMessage
    )
    
    try {
        $apiVersion = "7.0"
        $url = "$OrgUrl/$Project/_apis/git/repositories/$Repo/pushes?api-version=$apiVersion"
        
        # Get the repository ID and default branch
        $repoUrl = "$OrgUrl/$Project/_apis/git/repositories/$Repo?api-version=$apiVersion"
        $repoInfo = Invoke-RestMethod -Uri $repoUrl -Headers $Headers -Method Get
        
        # Get the latest commit on the default branch
        $refUrl = "$OrgUrl/$Project/_apis/git/repositories/$Repo/refs?filter=heads/$($repoInfo.defaultBranch)&api-version=$apiVersion"
        $ref = Invoke-RestMethod -Uri $refUrl -Headers $Headers -Method Get
        $oldObjectId = $ref.value[0].objectId
        
        # Prepare the changes
        $changes = @()
        
        # Get all files in the source directory
        $files = Get-ChildItem -Path $SourcePath -Recurse -File
        
        foreach ($file in $files) {
            $relativePath = $file.FullName.Replace($SourcePath, '').TrimStart('\')
            $targetPath = "$AppName/$Version/$relativePath"
            
            $content = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
            
            $changes += @{
                changeType = "add"
                item = @{
                    path = $targetPath
                }
                newContent = @{
                    content = $content
                    contentType = "base64"
                }
            }
        }
        
        # Prepare the push body
        $pushBody = @{
            refUpdates = @(
                @{
                    name = "refs/heads/$($repoInfo.defaultBranch)"
                    oldObjectId = $oldObjectId
                }
            )
            commits = @(
                @{
                    comment = $CommitMessage
                    changes = $changes
                }
            )
        } | ConvertTo-Json -Depth 20
        
        # Push the changes
        Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $pushBody
        $statusBox.AppendText("Successfully pushed all files in a single commit`r`n")
        return $true
    }
    catch {
        $statusBox.AppendText("Error during batch commit: $_`r`n")
        return $false
    }
}

# === Show Form ===
[void]$form.ShowDialog()
