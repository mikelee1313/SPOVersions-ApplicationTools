Here is a carefully crafted README file for your repository, based on the content and structure of Configure-SPOVersionsforAutomatic.ps1:

---

# Configure-SPOVersionsforAutomatic

This PowerShell script enables administrators to efficiently manage SharePoint Online site version policies and file version management across multiple sites, as defined in a text file. It provides a menu-driven interface for both querying and updating versioning policies, automating cleanup jobs, and monitoring storage usage.

## Features

- **Get current version policies** for all specified sites
- **Enable auto-expiration version trimming** to control version sprawl
- **Check version policy status and storage usage**
- **Create batch delete jobs** for version cleanup
- **Monitor batch deletion job status**
- Handles throttling and logs all operations for auditing and troubleshooting

## Prerequisites

- **PnP.PowerShell** module installed (Tested with version 3.1.0)
- **Site URLs file**: Create a text file listing each SharePoint Online site URL on a separate line at `C:\temp\M365CPI13246019-Sites.txt`
- Proper permissions to connect to SharePoint Online and modify site settings
- Microsoft 365 tenant and application (client) IDs
- Enable Tenant Level settings to Automatic

  <img width="1575" height="606" alt="image" src="https://github.com/user-attachments/assets/6458c57e-5d3d-4d43-a333-f8e6d6df2df6" />


## Setup

1. Install the PnP.PowerShell module if not already present:
   ```powershell
   Install-Module -Name "PnP.PowerShell" -Scope CurrentUser
   ```

2. Prepare the site list file:
   - Place the full URLs of all target SharePoint Online sites in `C:\temp\M365CPI13246019-Sites.txt`

3. Update the script, if necessary, to use your own Tenant ID, Client ID, and SharePoint Admin Center URL.

## Usage

Run the script from a PowerShell window:
```powershell
.\Configure-SPOVersionsforAutomatic.ps1
```

You will be presented with a menu to select the desired operation:

1. **Get current version policy for all sites**
2. **Set auto-expiration version trim to enabled for all sites**
3. **Get version policy status for all sites**
4. **Create batch delete job for all sites**
5. **Get batch delete job status for all sites**
6. **Q: Quit**

Operations are performed on all sites listed in your site file. The script will prompt for interactive authentication if needed.

## Logging and Output

- **Console Output:** Shows real-time status of operations per site.
- **Log Files:** Detailed logs are saved in your `%TEMP%` directory (e.g., `configure_versions_SPOyyyy-MM-dd_HH-mm-ss_logfile.log`).
- **Inputs:** Site URLs are read from the specified text file.
- **Outputs:** Operation statuses are reported on the console and logged.

## Example

```powershell
.\Configure-SPOVersionsforAutomatic.ps1
```
Follow the menu prompts to perform the desired operation across all specified sites.

## Authors

- Mike Lee
- Luis DuSolier

## Disclaimer

The sample script is provided **AS IS** without warranty of any kind. Microsoft and the authors disclaim all implied warranties, including but not limited to merchantability or fitness for a particular purpose. Use at your own risk.

---

Let me know if you want to further tailor this README (e.g., for company-specific branding or additional instructions)!
