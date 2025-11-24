# SharePoint Online Version Policy Management Script

A comprehensive PowerShell solution for managing SharePoint Online site version policies, batch deletion operations, and tenant-level version settings across your Microsoft 365 environment.

## 📋 Overview

This script provides an interactive menu-driven interface to manage SharePoint Online version policies at both site and tenant levels. It supports automatic (intelligent algorithm) and manual (configurable limits) version management modes, with robust batch processing capabilities for large-scale operations.

### Key Features

- **Site-Level Operations**: Configure version policies for individual sites or all sites in batch
- **Tenant-Level Operations**: Set default version policies for new sites
- **Batch Delete Jobs**: Clean up old file versions with flexible age-based or count-based deletion
- **Auto-Discovery Mode**: Automatically discover and process all SharePoint or OneDrive sites
- **File-Based Processing**: Process specific sites from a list file for targeted operations
- **Throttling Handling**: Built-in retry logic to handle SharePoint API throttling
- **Comprehensive Logging**: Detailed logging to troubleshoot and track all operations

## 🎯 Use Cases

- **Storage Optimization**: Reduce SharePoint storage consumption by removing old file versions
- **Compliance Management**: Enforce version retention policies across your tenant
- **Site Migration Preparation**: Clean up versions before migrating sites
- **Tenant Standardization**: Apply consistent version policies to all sites
- **OneDrive Management**: Manage version policies for all user OneDrive sites

## 📦 Prerequisites

### Required Software

- **PowerShell**: PowerShell 7.5
- **PnP.PowerShell Module**: Version 3.1.0 or later
  ```powershell
  Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
  ```

### Required Permissions

- **SharePoint Administrator** or **Global Administrator** role in Microsoft 365
- Permissions to connect to SharePoint Online admin center
- Permissions to modify site settings and create batch deletion jobs

### Azure AD App Registration (for authentication)

You'll need an Azure AD app registration with the following:
- **Application (client) ID**
- **Tenant ID**
- **API Permissions**:
  - `Sites.FullControl.All` (Application permission)
  - `Sites.ReadWrite.All` (Delegated permission)
  - `User.Read` (Delegated permission)

## ⚙️ Configuration

Before running the script, configure the following parameters in the script file:

### 1. Tenant Authentication Settings

```powershell
$tenantId = 'your-tenant-id'           # Your Microsoft 365 Tenant ID
$clientId = 'your-client-id'           # Azure AD Application (Client) ID
$url = "https://yourtenant-admin.sharepoint.com"  # SharePoint Admin Center URL
```

### 2. Site Discovery Mode

Choose between two modes:

#### Option A: File-Based Processing (Recommended for large tenants)
```powershell
$sitesFilePath = "C:\temp\Sites.txt"
```

Create a text file with one site URL per line:
```
https://contoso.sharepoint.com/sites/Site1
https://contoso.sharepoint.com/sites/Site2
https://contoso-my.sharepoint.com/personal/user_contoso_com
```

#### Option B: Auto-Discovery (Recommended for small tenants)
```powershell
$sitesFilePath = $null
```

When set to `$null`, the script will prompt you to select:
- **SharePoint sites** (excludes OneDrive and system sites)
- **OneDrive for Business sites** only

## 🚀 Getting Started

### Quick Start

1. **Clone or download** the script to your local machine
2. **Configure** the tenant settings and site discovery mode
3. **Run** the script in PowerShell:
   ```powershell
   .\Apply-SPOVersions-Tool.ps1
   ```
4. **Authenticate** when prompted using your SharePoint administrator credentials
5. **Select** an operation from the interactive menu

### First Run Example

```powershell
# Launch the script
.\Apply-SPOVersions-Tool.ps1

# You'll see a menu like this:
==== SharePoint Site Version Policy Operations ====

Site Mode: Auto-discovery (all sites in tenant)
  You will be prompted to select SharePoint or OneDrive sites

Site-Level Operations:
1: Get current version policy for all sites
2: Set version policy for all sites
3: Get version policy status for all sites
4: Create batch delete job for all sites
5: Get batch delete job status for all sites

Tenant-Level Operations (applies to new sites):
6: Set tenant to automatic version trimming
7: Set tenant to manual version limits
8: Review current tenant level version settings

Q: Quit
====================================================
Please select an operation (1-8, or Q to quit):
```

## 📖 Menu Options Explained

### Site-Level Operations (Options 1-5)

#### Option 1: Get Current Version Policy
**Purpose**: Retrieve and display the current version policy for all sites.

**Output**: Shows whether each site uses automatic trimming or manual limits.

**Use Case**: Audit current configuration before making changes.

---

#### Option 2: Set Version Policy
**Purpose**: Configure version policies for multiple sites at once.

**Sub-Options**:
- **Automatic Mode**: Uses Microsoft's intelligent algorithm to optimize storage
- **Manual Mode**: Set specific version limits and expiration rules
  - Use tenant-level settings (apply current tenant defaults)
  - Enter custom settings (specify your own limits)

**Parameters** (Manual Mode):
- **Major Version Limit**: Minimum 100 versions
- **Expire After Days**: 
  - Never (no expiration)
  - 90 days (3 months)
  - 180 days (6 months)
  - 365 days (1 year)
  - Custom (>29 days)

**Example Use Case**: Standardize all SharePoint sites to keep 500 versions with 180-day expiration.

---

#### Option 3: Get Version Policy Status
**Purpose**: Check the current status and storage impact of version policies.

**Output**: 
- Policy status (active/processing)
- Storage usage information
- Completion timestamps

**Use Case**: Monitor the effectiveness of version policies over time.

---

#### Option 4: Create Batch Delete Job
**Purpose**: Remove old file versions to free up storage space.

**Modes**:

##### Automatic Mode
- Uses each site's current version policy
- Respects tenant-level defaults
- Best for maintaining consistency

##### Manual Mode
Choose deletion method:

**Mode 1: Delete by Age**
- **Parameter**: `DeleteOlderThanDays`
- **Options**: 30, 90, 180, 365 days, or custom (≥30)
- **Behavior**: Removes all versions older than specified days
- **Example**: Delete all versions older than 90 days across all sites

**Mode 2: Delete by Count**
- **Parameter**: `MajorVersionLimit`
- **Options**: Any number ≥100
- **Behavior**: Keeps only the most recent X versions, deletes older ones
- **Example**: Keep only the 200 most recent versions per file

**Settings Source Options**:
- Use tenant-level settings (apply current defaults)
- Enter custom settings (specify your own values)

**Important**: 
- Batch delete jobs run asynchronously
- Check status with Option 5
- Cannot be undone once completed

---

#### Option 5: Get Batch Delete Job Status
**Purpose**: Monitor the progress of batch deletion operations.

**Output**:
- Job state (queued/processing/completed/failed)
- Completion timestamp
- Storage released (in bytes)
- Deletion mode used

**Use Case**: Track long-running batch delete operations and verify success.

---

### Tenant-Level Operations (Options 6-8)

#### Option 6: Set Tenant to Automatic Version Trimming
**Purpose**: Enable automatic version trimming for new sites.

**Effect**:
- New sites will use Microsoft's intelligent algorithm
- Existing sites retain their current settings
- Algorithm optimizes based on version creation date and access patterns

**Use Case**: Set organization-wide default to automatic mode.

---

#### Option 7: Set Tenant to Manual Version Limits
**Purpose**: Configure manual version limits as default for new sites.

**Parameters**:
- **Major Version Limit**: Minimum 100 versions
- **Expire After Days**: 
  - Never (no expiration)
  - 90 days (3 months)
  - 180 days (6 months)
  - 365 days (1 year)
  - Custom (>29 days minimum for tenant settings)

**Effect**:
- New sites inherit these settings
- Existing sites unchanged unless explicitly modified

**Use Case**: Enforce compliance policies requiring version retention limits.

---

#### Option 8: Review Current Tenant Settings
**Purpose**: Display current tenant-level version policy configuration.

**Output**:
- Current mode (automatic vs manual)
- Version limits (if manual mode)
- Expiration settings
- Explanation of impact on new vs existing sites

**Use Case**: Verify tenant configuration before bulk operations.

---

## 📝 Common Workflows

### Workflow 1: Standardize All Sites with Manual Policy

1. **Check tenant settings** (Option 8)
2. **Set tenant manual policy** (Option 7)
   - Example: 500 versions, 180 days
3. **Apply to all sites** (Option 2)
   - Choose "Manual mode"
   - Choose "Use tenant-level settings"
4. **Verify application** (Option 1)

### Workflow 2: Clean Up Storage Across OneDrive Sites

1. **Select auto-discovery mode** (set `$sitesFilePath = $null`)
2. **Create batch delete job** (Option 4)
   - Choose "Manual mode"
   - Choose "Delete by age"
   - Select 90 days
   - When prompted, select "OneDrive for Business sites"
3. **Monitor progress** (Option 5)
4. **Check results** after completion

### Workflow 3: Migrate from Manual to Automatic Mode

1. **Review current settings** (Option 8)
2. **Set tenant to automatic** (Option 6)
3. **Apply to existing sites** (Option 2)
   - Choose "Automatic mode"
4. **Verify policy status** (Option 3)

### Workflow 4: Process Specific Sites from List

1. **Create site list file** (e.g., `C:\temp\Sites.txt`)
2. **Configure script** (`$sitesFilePath = "C:\temp\Sites.txt"`)
3. **Run desired operation** (Options 1-5)
4. **Review log file** for detailed results

## 🔧 Troubleshooting

### Authentication Issues

**Problem**: Repeated login prompts during batch processing

**Solution**: This was fixed in the latest version. Ensure you're using the script without the `-Connection` parameter in `Invoke-SiteBatch`.

---

### Throttling Errors (HTTP 429 or 503)

**Problem**: "Too many requests" errors

**Solution**: Built-in throttling handler automatically retries with exponential backoff (up to 5 attempts). If persistent:
- Reduce batch size (use file-based processing with smaller file)
- Run during off-peak hours
- Check Microsoft 365 service health

---

### Parameter Validation Errors

**Problem**: "Cannot bind parameter" or "Parameter set cannot be resolved"

**Solutions**:

- **Set-PnPSiteVersionPolicy errors**:
  - Requires ALL three parameters: `MajorVersions`, `MajorWithMinorVersions`, `ExpireVersionsAfterDays`
  - Use `ExpireVersionsAfterDays = 0` for "Never"

- **New-PnPSiteFileVersionBatchDeleteJob errors**:
  - Mode 1 (age): Use ONLY `DeleteOlderThanDays`
  - Mode 2 (count): Use ONLY `MajorVersionLimit` + `MajorWithMinorVersionsLimit`
  - Cannot mix parameters from different modes

---

### Site Discovery Issues

**Problem**: No sites found or discovery fails

**Solutions**:
- Verify you have SharePoint Administrator permissions
- Check tenant URL is correct (admin center URL)
- For OneDrive sites, ensure user profiles are provisioned
- Review template filters in `Get-FilteredSites` function

---

### Log File Location

All operations are logged to: `%TEMP%\configure_versions_SPO[date]_logfile.log`

**Example**: `C:\Users\YourName\AppData\Local\Temp\configure_versions_SPO2025-11-24_14-30-00_logfile.log`

**Log Levels**:
- **INFO**: Normal operations and confirmations
- **WARNING**: Non-critical issues (e.g., throttling, retries)
- **ERROR**: Failures requiring attention
- **DEBUG**: Detailed diagnostic information (enabled when `$Debug = $true`)

## 🔍 Technical Details

### API Cmdlets Used

- `Connect-PnPOnline`: Authenticates to SharePoint
- `Get-PnPTenant`: Retrieves tenant-level settings
- `Set-PnPTenant`: Configures tenant-level policies
- `Get-PnPTenantSite`: Discovers sites in tenant
- `Get-PnPSiteVersionPolicy`: Gets site version policy
- `Set-PnPSiteVersionPolicy`: Sets site version policy
- `Get-PnPSiteVersionPolicyStatus`: Checks policy status
- `New-PnPSiteFileVersionBatchDeleteJob`: Creates deletion jobs
- `Get-PnPSiteFileVersionBatchDeleteJobStatus`: Monitors deletion jobs

### Template Exclusions (Auto-Discovery)

The following site templates are excluded from SharePoint site discovery:
- `RedirectSite#0`: Redirect sites
- `SRCHCEN*`: Search center sites
- `SRCHCENTERLITE*`: Search center lite sites
- `SPSMSITEHOST*`: SharePoint site host sites
- `APPCATALOG*`: App catalog sites
- `REDIRECTSITE*`: Redirect site templates
- Sites with `-my.sharepoint.com/personal/*` in URL (OneDrive sites)

### Script Scope Variables

The script uses `$script:` scope for variables accessed within scriptblocks:
- `$script:currentMajorVersionLimit`: Version count for manual policy
- `$script:currentExpireAfterDays`: Expiration days for manual policy
- `$script:currentDeleteOlderThanDays`: Age threshold for batch delete
- `$script:currentDeleteMajorVersionLimit`: Count limit for batch delete

### Error Handling

- **Throttling**: Automatic retry with exponential backoff (30, 60, 120, 240, 480 seconds)
- **Connection failures**: Logged with full exception details
- **Validation errors**: User-friendly messages with guidance
- **Site-level failures**: Continue processing remaining sites

## ⚠️ Important Considerations

### Version Deletion is Permanent

- Batch delete jobs **cannot be undone**
- Deleted versions are permanently removed from the recycle bin
- Test on a small set of sites before large-scale operations
- Consider backup/retention policies before deletion

### Tenant vs Site Settings

- **Tenant settings** apply to NEW sites only
- **Existing sites** require explicit modification (Option 2)
- Sites can override tenant defaults with custom policies

### Performance Expectations

- **Small tenants** (<100 sites): 5-10 minutes per operation
- **Medium tenants** (100-1000 sites): 30-60 minutes per operation
- **Large tenants** (>1000 sites): Several hours, recommend file-based processing
- **Batch delete jobs**: Run asynchronously, completion time varies by version count

### API Limits

- **Throttling threshold**: ~600 requests per minute per app
- **Automatic retry**: Up to 5 attempts with exponential backoff
- **Recommendation**: Process sites in batches during off-peak hours for large tenants

## 📊 Version Policy Comparison

| Feature | Automatic Mode | Manual Mode |
|---------|---------------|-------------|
| **Storage Optimization** | Intelligent algorithm | Fixed limits |
| **Configuration** | No settings required | Define limits + expiration |
| **Predictability** | Algorithm-based | Exact version count/age |
| **Compliance** | Varies by algorithm | Precise control |
| **Best For** | General use, storage optimization | Compliance requirements, predictable retention |
| **Version Limit** | N/A | Minimum 100 |
| **Expiration** | Algorithm-determined | Never or >29 days |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

### Reporting Issues

When reporting issues, please include:
- PowerShell version (`$PSVersionTable`)
- PnP.PowerShell module version (`Get-Module PnP.PowerShell`)
- Error messages from console
- Relevant log file excerpts
- Steps to reproduce

## 📄 License

This sample script is provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.

## 👥 Authors

- **Mike Lee** - Initial development
- **Luis DuSolier** - Contributions and enhancements

## 📅 Version History

- **v1.0.0** (November 2025) - Initial release
  - Site-level and tenant-level version policy management
  - Batch delete operations with automatic and manual modes
  - Auto-discovery and file-based site processing
  - Comprehensive logging and error handling
  - Throttling retry logic

## 🔗 Additional Resources

- [Microsoft Documentation: File versioning in SharePoint](https://learn.microsoft.com/en-us/sharepoint/file-versioning)
- [PnP PowerShell Documentation](https://pnp.github.io/powershell/)
- [SharePoint Storage Management](https://learn.microsoft.com/en-us/sharepoint/manage-site-collection-storage-limits)
- [Azure AD App Registration Guide](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

## 📧 Support

For issues or questions:
1. Check the [Troubleshooting](#-troubleshooting) section
2. Review the log file in `%TEMP%`
3. Open an issue on GitHub with detailed information

---

**Note**: This script requires appropriate SharePoint administrator permissions and should be tested in a non-production environment before deployment to production.
