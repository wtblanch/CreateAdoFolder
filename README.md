# Azure DevOps File Uploader

A PowerShell-based GUI application that simplifies the process of uploading files to Azure DevOps repositories.

## Features

- User-friendly graphical interface
- Azure DevOps integration
- SharePoint URL support (up to 3 URLs)
- File and folder upload capabilities
- Real-time status updates
- Configurable connection settings

## Requirements

- Windows operating system
- PowerShell
- Azure DevOps access
- Required PowerShell modules:
  - System.Windows.Forms
  - System.Drawing

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

2. Fill in the required connection and project settings in the GUI
3. Select files or folders to upload
4. Click the Start button to begin the upload process
5. Monitor the status in the status box

## Interface

The application provides a clean interface with the following sections:
- Connection Settings
- Project Settings
- File Selection
- Status Updates

