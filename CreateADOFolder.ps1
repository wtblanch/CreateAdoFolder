Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
$form.Size = New-Object System.Drawing.Size(600, 800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize = New-Object System.Drawing.Size(600, 800)
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
$spTextBox.Size = New-Object System.Drawing.Size(500, 70)
$spTextBox.Location = New-Object System.Drawing.Point(20, 50)
$spTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
# Pre-fill SharePoint URLs from config
if ($config.SharePointUrls) {
    $spTextBox.Text = $config.SharePointUrls -join "`r`n"
}
$spGroup.Controls.Add($spTextBox)

# === File Picker Button ===
$localFiles = @()
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Add Local Files"
$browseButton.Location = New-Object System.Drawing.Point(20, 540)
$browseButton.Size = New-Object System.Drawing.Size(540, 30)
$browseButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$browseButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$browseButton.ForeColor = [System.Drawing.Color]::White
$browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$browseButton.Add_Click({
    $openFile = New-Object System.Windows.Forms.OpenFileDialog
    $openFile.Multiselect = $true
    $openFile.Filter = "All Files (*.*)|*.*"
    if ($openFile.ShowDialog() -eq "OK") {
        $localFiles = $openFile.FileNames
        $statusBox.AppendText("Selected local files: `r`n" + ($localFiles -join "`r`n") + "`r`n")
    }
})
$form.Controls.Add($browseButton)

# === Status Box ===
$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.Size = New-Object System.Drawing.Size(540, 60)
$statusBox.Location = New-Object System.Drawing.Point(20, 580)
$statusBox.ReadOnly = $true
$statusBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$statusBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($statusBox)

# === Start Button ===
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start Upload"
$startButton.Location = New-Object System.Drawing.Point(20, 650)
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

    # === Function to create file in Azure DevOps ===
    function Add-FileToADO {
        param (
            $FilePath,
            $Content,
            $CommitMessage
        )

        $apiVersion = "7.0"
        $encodedContent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Content))
        
        $body = @{
            content = $encodedContent
            contentType = "base64"
            message = $CommitMessage
        } | ConvertTo-Json

        $url = "$orgUrl/$project/_apis/git/repositories/$repo/pushes?api-version=$apiVersion"
        
        try {
            Invoke-RestMethod -Uri $url -Headers $headers -Method Post -Body $body
            $statusBox.AppendText("Successfully created: $FilePath`r`n")
        }
        catch {
            $statusBox.AppendText("Error creating $FilePath : $_`r`n")
        }
    }

    try {
        # === Generate README.md ===
        $readmeContent = @"
# $($textBoxes["App Name"].Text)

## Version: $($textBoxes["Version"].Text)

This repository contains the files and configuration for $($textBoxes["App Name"].Text).

## Repository Information
- **Organization:** $($textBoxes["Organization URL"].Text)
- **Project:** $($textBoxes["Project"].Text)
- **Repository:** $($textBoxes["Repo"].Text)

## Contents
This folder contains:
- Configuration files
- SharePoint document references
- Application-specific settings

## SharePoint Documents
The following SharePoint documents are referenced:
$(($spTextBox.Text -split "`r`n" | Where-Object { $_ -ne "" } | ForEach-Object { "- $_" }) -join "`r`n")

## Version History
- Current Version: $($textBoxes["Version"].Text)
  - Initial setup and configuration
  - Basic file structure established
  - SharePoint documents linked

## Contact
For questions or issues, please contact the repository maintainers.

---
Last Updated: $(Get-Date -Format "yyyy-MM-dd")
"@

        Add-FileToADO -FilePath "$($textBoxes["App Name"].Text)/$($textBoxes["Version"].Text)/README.md" `
                      -Content $readmeContent `
                      -CommitMessage "Added README.md"

        # === Process SharePoint URLs ===
        $urls = $spTextBox.Text -split "`r`n" | Where-Object { $_ -ne "" }
        
        if ($urls.Count -gt 0) {
            $urlContent = $urls | ConvertTo-Json
            Add-FileToADO -FilePath "$appName/$version/urls.json" -Content $urlContent -CommitMessage "Added SharePoint URLs"
        }

        # === Process local files ===
        foreach ($file in $localFiles) {
            $fileName = Split-Path $file -Leaf
            $content = Get-Content $file -Raw
            Add-FileToADO -FilePath "$appName/$version/$fileName" -Content $content -CommitMessage "Added $fileName"
        }

        $statusBox.AppendText("Upload completed successfully.`r`n")
    }
    catch {
        $statusBox.AppendText("Error during upload: $_`r`n")
    }
})
$form.Controls.Add($startButton)

# === Show Form ===
[void]$form.ShowDialog()
