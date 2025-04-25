# Azure DevOps File Uploader

A PowerShell-based GUI application that simplifies the process of uploading files to Azure DevOps repositories.

## Features

- User-friendly graphical interface
- Azure DevOps integration
- SharePoint URL validation and file synchronization
- File and folder selection with visual list
- Batch upload in single commit
- Real-time status updates
- Configurable connection settings

## Requirements

- Windows operating system
- PowerShell
- Azure DevOps access
- Required PowerShell modules:
  - System.Windows.Forms
  - System.Drawing
  - PnP.PowerShell (for SharePoint integration)

## Installation

1. Install required PowerShell module:

   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser
   ```

## Configuration

Edit the configuration section in `CreateADOFolder.ps1`:

```powershell
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
```

## Usage

1. Run the script:

   ```powershell
   .\CreateADOFolder.ps1
   ```

2. Fill in the required connection settings:
   - Personal Access Token (PAT)
   - Organization URL
   - Project name
   - Repository name
   - Application name
   - Version

3. Add SharePoint URLs:
   - Enter SharePoint URLs in the text box (one per line)
   - Click "Validate URLs" to verify format
   - URLs must be in the format: https://*.sharepoint.com/*
   - ** for use later

4. Select files and folders:
   - Click "Add Files" to select individual files
   - Click "Add Folder" to select entire folders
   - Selected items appear in the list view
   - Use "Remove Selected" to remove items from the list

5. Click "Start Upload" to begin the process:
   - Files are downloaded from SharePoint
   - Local files are collected
   - README.md is generated
   - All files are uploaded in a single commit

## Repository Structure

When files are uploaded, they follow this structure:

```
AppName/
└── Version/
    ├── README.md
    ├── [SharePoint Files with Original Structure]
    │   ├── folder1/
    │   │   ├── file1.docx
    │   │   └── file2.xlsx
    │   └── folder2/
    │       └── file3.pdf
    └── [Local Files]
        └── file4.txt
```

## Interface

The application provides a clean interface with the following sections:

- Connection Settings
  - PAT and Organization URL inputs
- Project Settings
  - Project, Repo, App Name, Version inputs
- SharePoint URLs
  - URL input with validation
- Local Files and Folders
  - File and folder selection
  - List view of selected items
- Status Updates
  - Real-time progress and error messages
