# Apply-SPOVersions-Tool

A PowerShell script for managing SharePoint Online file version policies across multiple site collections at scale. Built on [PnP PowerShell](https://pnp.github.io/powershell/), it provides an interactive menu to audit version policies, apply automatic or manual version limits, run batch delete jobs, generate version history reports, and perform What-If storage impact analysis — all across an entire tenant or a defined list of site collections.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [App Registration Setup](#app-registration-setup)
- [Configuration](#configuration)
- [Site Discovery Modes](#site-discovery-modes)
- [Running the Script](#running-the-script)
- [Menu Options](#menu-options)
  - [Site-Level Operations](#site-level-operations)
  - [Tenant-Level Operations](#tenant-level-operations)
- [Version Report Library](#version-report-library)
- [What-If Analysis](#what-if-analysis)
- [Logging](#logging)
- [Output Files](#output-files)
- [Throttling Handling](#throttling-handling)
- [Disclaimer](#disclaimer)

---

## Overview

SharePoint Online version history can consume significant storage over time. This tool helps administrators:

1. **Audit** existing version policies across all site collections.
2. **Enforce** automatic or manual version limits at the site and tenant level.
3. **Clean up** excess versions using batch delete jobs.
4. **Report** on version storage usage per site collection.
5. **Analyze** projected storage savings before committing to a policy change using What-If analysis.

The script is entirely interactive — no parameters need to be passed at the command line. All configuration is done inside the script file itself.

---

## Prerequisites

| Requirement | Details |
|---|---|
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ (recommended for best compatibility) |
| **PnP.PowerShell** | Version 3.1.0 or later. Install with: `Install-Module PnP.PowerShell` |
| **Microsoft 365 Tenant** | SharePoint Online with admin access |
| **Entra ID App Registration** | Required for interactive OAuth authentication (see below) |
| **SharePoint Administrator role** | Required to set tenant/site version policies and generate reports |
| **Sites.FullControl.All** | Delegated API permission required on the app registration |

---

## App Registration Setup

The script authenticates using an Entra ID (Azure AD) app registration with delegated permissions via interactive browser login.

### Steps

1. In the [Azure Portal](https://portal.azure.com), go to **Entra ID → App registrations → New registration**.
2. Set a name (e.g., `SPO-VersionManagement-Tool`) and select **Single tenant**.
3. Under **Authentication**, add a **Mobile and desktop application** redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`.
4. Enable **Allow public client flows**.
5. Under **API permissions**, add the following **Delegated** permissions and grant admin consent:

| API | Permission | Type | Required For |
|---|---|---|---|
| SharePoint | `AllSites.FullControl` | Delegated | All operations |
| SharePoint | `AllSites.Manage` | Delegated | Version policy management |

6. Copy the **Application (client) ID** and **Directory (tenant) ID** into the script configuration section.

> **Note:** `Sites.FullControl.All` as a **delegated** permission is specifically required to generate version history report jobs (`New-PnPSiteFileVersionExpirationReportJob`). Delegated `Sites.Manage.All` alone is insufficient for report generation.

---

## Configuration

Edit the configuration section at the top of the script:

```powershell
#tenant Properties
$tenantId  = '<your-tenant-id>'
$clientId  = '<your-app-registration-client-id>'
$url       = "https://<your-tenant>-admin.sharepoint.com"
```

### Site List Configuration

```powershell
# Option 1: Load from file (recommended for large tenants)
$sitesFilePath = "C:\temp\MySites.txt"

# Option 2: Auto-discover all sites at runtime
$sitesFilePath = $null
```

**Option 1 — Sites file:** Create a plain text file with one site URL per line:

```
https://contoso.sharepoint.com
https://contoso.sharepoint.com/sites/Finance
https://contoso.sharepoint.com/sites/IT
```

**Option 2 — Auto-discovery:** Set `$sitesFilePath = $null` and the script will prompt you to choose between SharePoint sites or OneDrive sites at runtime. System sites (search centres, app catalogs, redirect sites) are automatically excluded.

### Version Report Scope Filter

The script includes a dedicated size threshold for version report operations:

```powershell
# StorageUsageCurrent is in MB. Set to 0 to include all sites.
$MinSiteSizeforversionReports = 100
```

This filter uses each site's `StorageUsageCurrent` value from tenant site metadata and applies only to:

- **Option 9** (Generate version history report)
- **Option 10** (Get version history report job status)
- **Option 11** (What-If analysis)

All other operations (for example version policy set/apply and batch delete jobs) still run against the full selected site scope.

At runtime, options 9, 10, and 11 also display the active threshold before execution so you can confirm the current scope.

---

## Site Discovery Modes

When auto-discovery is enabled (`$sitesFilePath = $null`), the script discovers sites at operation time:

- **SharePoint sites** — All sites excluding OneDrive personal sites and system templates (SRCHCEN, APPCATALOG, SPSMSITEHOST, REDIRECTSITE).
- **OneDrive for Business** — Only personal OneDrive sites (`-my.sharepoint.com/personal/`).

You will be asked to confirm the number of sites before any operation proceeds.

---

## Running the Script

```powershell
.\Apply-SPOVersions-Tool.ps1
```

On first run, an interactive browser window will open for authentication. Subsequent connections within the same session reuse the token silently. A new browser prompt appears when switching between site connections during batch operations.

---

## Menu Options

```
==== SharePoint Site Version Policy Operations ====

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
9: Generate version history report for all sites
10: Get version history report job status for all sites
11: What-If analysis - estimate storage recovery by version policy

Q: Quit
```

---

### Site-Level Operations

These operations connect to **each site collection individually** and use the existing throttling-aware batch processing framework.

#### Option 1 — Get Current Version Policy

Retrieves the current version policy (`Get-PnPSiteVersionPolicy`) for every site collection. Output shows whether each site is configured for automatic trimming or manual limits, and the specific limit values.

#### Option 2 — Set Version Policy for All Sites

Applies a version policy to every site collection using `Set-PnPSiteVersionPolicy`. At prompt you choose:

- **Automatic** — SharePoint uses an intelligent algorithm to expire older versions based on age. No explicit version count is needed.
- **Manual** — You specify:
  - **Major version limit** (minimum 100) — maximum number of major versions to retain.
  - **Expire after days** — optionally expire versions older than a specified number of days (minimum 30), or never.

  Settings can be sourced from current **tenant-level defaults** or entered as custom values.

The following image depicts the restore options and the storage use for each setting:
<img width="3748" height="1845" alt="image" src="https://github.com/user-attachments/assets/e9efce2e-6516-4f2f-b366-9745fd6238be" />


#### Option 3 — Get Version Policy Status

Retrieves the current policy propagation status (`Get-PnPSiteVersionPolicyStatus`) for each site, showing whether any pending policy changes have been applied.

#### Option 4 — Create Batch Delete Job

Submits a batch delete job (`New-PnPSiteFileVersionBatchDeleteJob`) on each site to delete excess versions according to the chosen policy. At prompt you choose:

- **Automatic** — Deletes versions that fall outside the site's current automatic policy.
- **Manual** — Choose one deletion criteria:
  - `DeleteOlderThanDays` — Remove all versions older than N days (minimum 30).
  - `MajorVersionLimit` — Keep only the N most recent major versions (minimum 100), with `MajorWithMinorVersionsLimit = 0`.
  - `DeleteBeforeDays` — Delete all versions created before N days ago.

  Settings can be pulled from **tenant-level defaults** or entered as custom values.

> **Warning:** Batch delete is a destructive operation. Deleted versions cannot be recovered from the Recycle Bin.

#### Option 5 — Get Batch Delete Job Status

Retrieves the status of any active or recently completed batch delete job (`Get-PnPSiteFileVersionBatchDeleteJobStatus`) per site, showing state, completion time, delete mode, and storage released in bytes.

---

### Tenant-Level Operations

These operations apply to **new site collections** created after the setting is applied. Existing sites are not affected unless explicitly updated via Option 2.

#### Option 6 — Set Tenant to Automatic Version Trimming

Enables automatic version trimming at the tenant level (`Set-PnPTenant -EnableAutoExpirationVersionTrim $true`). New sites will automatically use SharePoint's intelligent age-based expiration algorithm.

#### Option 7 — Set Tenant to Manual Version Limits

Configures the tenant to use manual version limits (`Set-PnPTenant -EnableAutoExpirationVersionTrim $false`). You specify:

- **Major version limit** (minimum 100).
- **Expire after days** — set to 0 for "Never", or a value of 30+ days.

> Note: The tenant-level minimum for `ExpireVersionsAfterDays` is 30 days. Values below 30 are automatically adjusted.

#### Option 8 — Review Current Tenant Version Settings

Displays current tenant version policy settings (`Get-PnPTenant`) including the active mode (Automatic vs. Manual), major version limit, and expiration days.

#### Option 9 — Generate Version History Report for All Sites

Submits an asynchronous version history report job (`New-PnPSiteFileVersionExpirationReportJob`) for each in-scope site collection. The in-scope list is filtered by `$MinSiteSizeforversionReports` (MB) using `StorageUsageCurrent`. The report is a CSV file saved to a **dedicated document library** created automatically in each site:

```
Admin_SiteCollection_VersionReport_DONOTDELETE
```

The library is created if it does not already exist (using `New-PnPList -Template DocumentLibrary`). This dedicated library avoids the metadata conflict issues that can occur when reports are written to Shared Documents.

**Report filename convention:**

```
{SiteCollectionName}site_adminreport_donotdelete_VersionReport.csv
```

For example, for `https://contoso.sharepoint.com/sites/Finance`:

```
Financesite_adminreport_donotdelete_VersionReport.csv
```

> Report generation is asynchronous. The job is submitted and runs in the background. Use **Option 10** to check when it has completed before running **Option 11**.

Before the job starts, the script prints the active threshold:

`Version report size threshold (MinSiteSizeforversionReports): <value> MB`

#### Option 10 — Get Version History Report Job Status

Checks the status of the report generation job (`Get-PnPSiteFileVersionExpirationReportJobStatus`) for each in-scope site using the same library and filename convention as Option 9.

The same `$MinSiteSizeforversionReports` filter is applied so status checks run against the same site scope used for report generation.

Displays per-site status (`completed`, `failed`, or in-progress) and at the end shows a summary:

```
==== Report Status Summary ====
  Total sites processed : 13
  Completed             : 11
  Failed                : 2

  Failed sites:
    - https://contoso.sharepoint.com/sites/IT
      The operation cannot continue because the report file has been modified...
```

**Common failure causes:**

- The report CSV was opened in Excel or another application while the job was still writing.
- Re-run Option 9 for any failed sites once the file is closed.

#### Option 11 — What-If Analysis

Downloads the completed CSV reports from each site and calculates the projected storage savings if a specific version policy were applied — **without making any changes**. Based on the [Microsoft What-If tutorial](https://learn.microsoft.com/en-us/sharepoint/tutorial-run-what-if-analysis).

The same `$MinSiteSizeforversionReports` filter is applied first, so analysis runs only for the same in-scope sites used by options 9 and 10.

See the [What-If Analysis](#what-if-analysis) section for full details.

---

## Version Report Library

Reports are stored in a dedicated document library created in each site collection:

| Property | Value |
|---|---|
| **Library title** | `Admin_SiteCollection_VersionReport_DONOTDELETE` |
| **Library URL** | `{siteUrl}/Admin_SiteCollection_VersionReport_DONOTDELETE/` |
| **Template** | Document Library |
| **Created by** | Option 9 (automatically, if not present) |

> Do **not** rename, move, or modify the CSV files while a report job is in progress. Doing so causes the job to fail with a "file modified" error.

---

## What-If Analysis

**Option 11** performs a read-only storage impact analysis using the version report CSVs. It downloads each CSV locally, applies a simulated policy in-memory, and reports how much storage would be recovered.

### How It Works

1. Prompts you to select a version policy to simulate.
2. Applies the `$MinSiteSizeforversionReports` filter using `StorageUsageCurrent`.
3. Downloads the report CSV from each in-scope site's `Admin_SiteCollection_VersionReport_DONOTDELETE` library to a local temp folder.
4. Expands the "compact" CSV format (where repeated field values are omitted for efficiency).
5. Applies the chosen algorithm to set `TargetExpirationDate` on each version row.
6. Sums the `Size` field for all rows that have a `TargetExpirationDate` set (versions that would be deleted).
7. Displays per-site and aggregate results.
8. Exports results to a CSV file for further analysis.
9. Optionally cleans up the downloaded temp files.

### Policy Options

| Option | Algorithm | Parameters |
|---|---|---|
| **Automatic** | Copies `AutomaticPolicyExpirationDate` from the report into `TargetExpirationDate` | None |
| **Expire After Days** | Sets `TargetExpirationDate = SnapshotDate + N days` | Days (minimum 30) |
| **Count Limit** | Keeps the N most recent major versions per file; marks all others with expiration `2000-01-01` | Major version count (minimum 1) |

### Sample Output

```
==== What-If Analysis Summary ====
  Policy analyzed         : Manual count limit: keep 100 most recent major versions
  Sites analyzed          : 13
  Total versions          : 485,320
  Versions to delete      : 312,100
  Total version storage   : 8,450 MB
  Total storage to recover: 5,230 MB  (5.107 GB)
  Overall % recovered     : 61.9%

  Per-site breakdown (sorted by storage freed):
    https://.../sites/ProductionDepartment         1,820 MB freed  (74%)
    https://.../sites/SalesandMarketing              950 MB freed  (68%)
    ...

  Results exported to: C:\Users\...\AppData\Local\Temp\SPO_WhatIf_Results_...csv
```

### Exported CSV Columns

| Column | Description |
|---|---|
| `SiteUrl` | Site collection URL |
| `TotalVersions` | Total version rows in the report |
| `VersionsToDelete` | Versions that would be deleted under the policy |
| `TotalVersionStorageMB` | Total storage used by all versions (MB) |
| `StorageFreedMB` | Storage that would be recovered (MB) |
| `StorageFreedGB` | Storage that would be recovered (GB) |
| `PercentFreed` | Percentage of version storage that would be freed |

The last row is a **TOTAL** row summarising all sites.

> **Prerequisite:** Option 9 must have been run and the report jobs must be **completed** (verified via Option 10) before running Option 11. Sites without a completed report CSV are skipped.

---

## Logging

Every script run creates a timestamped log file in the user's `%TEMP%` directory:

```
%TEMP%\configure_versions_SPO<yyyy-MM-dd_HH-mm-ss>_logfile.log
```

Log entries include timestamp, log level (`INFO`, `WARNING`, `ERROR`, `DEBUG`), and descriptive message. Debug-level entries are written when `$Debug = $true` (the default).

---

## Output Files

| File | Location | Created By |
|---|---|---|
| **Session log** | `%TEMP%\configure_versions_SPO*.log` | Every script run |
| **Version report CSV** | `{siteUrl}/Admin_SiteCollection_VersionReport_DONOTDELETE/{name}.csv` | Option 9 |
| **What-If results CSV** | `%TEMP%\SPO_WhatIf_Results_*.csv` | Option 11 |
| **Downloaded report CSVs** | `%TEMP%\SPO_WhatIf_<timestamp>\` | Option 11 (optionally kept) |

---

## Throttling Handling

All site-level batch operations go through `Invoke-WithThrottlingHandling`, which:

- Catches HTTP 429 (Too Many Requests) and HTTP 503 (Service Unavailable) responses.
- Reads the `Retry-After` header and waits the specified number of seconds before retrying.
- Falls back to exponential back-off (`InitialRetrySeconds × 2^retryCount`) if no `Retry-After` header is present.
- Retries up to **5 times** per site before logging a failure and moving on.

---

## Disclaimer

> The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

---

## Authors

- **Mike Lee**
- **Luis DuSolier**

*Script created: November 2025 | Last updated: July 2026*
