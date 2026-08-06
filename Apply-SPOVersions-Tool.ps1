<#
.SYNOPSIS
    PowerShell script to manage SharePoint Online site version policies across multiple sites.

.DESCRIPTION
    This script provides functionality to manage SharePoint Online site version policies and file version management across multiple sites defined in a text file. 
    
    It includes capabilities to:

    - Get current version policies
    - Enable auto-expiration version trimming
    - Check version policy status and storage usage
    - Create batch delete jobs for version cleanup
    - Monitor batch deletion job status

    The script implements throttling handling to manage SharePoint Online request limits and provides detailed logging.

.PARAMETER tenantId
    The Microsoft 365 tenant ID.

.PARAMETER clientId
    The application (client) ID for authentication.

.PARAMETER url
    The SharePoint Online admin center URL.

.EXAMPLE
    .\Apply-SPOVersions-Tool.ps1

.NOTES
    Authors: Mike Lee / Luis DuSolier / Joseph Vasil
    Date: 11/24/25
    Updated: 7/31/26 - added Version report generation and  What-If analysis and CSV export functionality
    Updated: 8/3/26 - Fixed bug where checking version report was calling the library and not site, fixed 3 functions to use throttle handling.
    Updated: 8/6/26 - Added site size filter using $MinSiteSizeforversionReports for version report generation and What-If analysis, added logging for 
        skipped sites due to size filter, added logging for sites included without size metadata.

    File Name      : Apply-SPOVersions-Tool.ps1
    Prerequisites  : 
    - PnP.PowerShell module installed (Tested with 3.1.0)
    - Text file with site URLs at C:\temp\M365CPI13246019-Sites.txt
    - Proper permissions to connect to SPO and modify sites
    
    The script uses interactive authentication. Make sure you have appropriate permissions
    to perform operations on the specified SharePoint sites.

.Disclaimer: The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

.INPUTS
    Site URLs from a text file located at C:\temp\M365CPI13246019-Sites.txt.

.OUTPUTS
    - Console output showing operation status
    - Detailed log file in %TEMP% directory named 'configure_versions_SPO[date]_logfile.log'

.FUNCTIONALITY
    SharePoint Online, Version Management, Site Management, PnP PowerShell
#>

# Initialize logging
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$log = "$env:TEMP\" + 'configure_versions_SPO' + $date + '_' + "logfile.log"
$EnableDebugLogging = $true

# Require PowerShell 5.1 or later
if ($PSVersionTable.PSVersion.Major -lt 5 -or
    ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
    Write-Host "ERROR: This script requires PowerShell 5.1 or later." -ForegroundColor Red
    Write-Host "  Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    exit
}

# Verify PnP.PowerShell is installed
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    Write-Host "ERROR: The PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "  Install it by running:  Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
    exit
}

# This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"  # Default log level is INFO
    )
    if ($null -ne $LogName) {
        # Skip DEBUG level messages when debug logging is disabled
        if ($LogLevel -eq "DEBUG" -and $EnableDebugLogging -eq $false) {
            return
        }
        
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : [$LogLevel] $LogEntryText" | Out-File -FilePath $LogName -append;
    }
}

############################################
################configuration###############

#tenant Properties
$tenantId = '9cfc42cb-51da-4055-87e9-b20a170b6ba3'
$clientId = '1e892341-f9cd-4c54-82d6-0fc3287954cf'
$url = "https://m365cpi13246019-admin.sharepoint.com"

# Site Discovery Configuration
# =============================
# Option 1: Process specific sites from a file (recommended for large tenants)
#   - Set $sitesFilePath to the path of a text file containing site URLs (one per line)
#   - Example: $sitesFilePath = "C:\temp\M365CPI13246019-Sites.txt"
#
# Option 2: Process ALL sites in the tenant automatically (recommended for small tenants)
#   - Set $sitesFilePath = $null
#   - Script will prompt to choose between SharePoint sites or OneDrive sites
#   - SharePoint sites exclude system sites (search centers, app catalog, etc.)
#   - OneDrive sites target personal sites only

#$sitesFilePath = "C:\temp\M365CPI13246019-Sites.txt"  # Set to $null to auto-discover all sites
$sitesFilePath = $null # Set to $null to auto-discover all sites

# Version report scope configuration
# StorageUsageCurrent is reported in MB. Set to 0 to include all sites in report generation and What-If analysis.
$MinSiteSizeforversionReports = 100

#################section####################
############################################

# Function to get site scope from user when auto-discovering sites
function Get-SiteScope {
    Write-Host "`n==== Select Site Scope for Auto-Discovery ====" -ForegroundColor Cyan
    Write-Host "1: SharePoint sites (excludes OneDrive and system sites)"
    Write-Host "2: OneDrive for Business sites only"
    Write-Host "3: Cancel and return to menu"
    
    $scopeChoice = $null
    do {
        $scopeChoice = Read-Host "Select site scope (1-3)"
        if ($scopeChoice -notin @("1", "2", "3")) {
            Write-Host "Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red
        }
    } while ($scopeChoice -notin @("1", "2", "3"))
    
    return $scopeChoice
}

# Function to discover and filter sites based on scope
function Get-FilteredSites {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Scope
    )
    
    Write-Host "`nDiscovering sites in tenant..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting site discovery with scope: $Scope" -LogLevel "INFO"
    
    try {
        if ($Scope -eq "1") {
            # Get all SharePoint sites, excluding OneDrive and system sites
            Write-Host "Retrieving SharePoint sites (excluding OneDrive and system sites)..." -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Retrieving SharePoint sites with template filters" -LogLevel "INFO"
            
            $allSites = Get-PnPTenantSite | Where-Object {
                $_.Template -ne 'RedirectSite#0' -and
                $_.ArchiveStatus -eq "NotArchived" -and
                $_.Template -notlike 'SRCHCEN*' -and
                $_.Template -notlike 'SRCHCENTERLITE*' -and
                $_.Template -notlike 'SPSMSITEHOST*' -and
                $_.Template -notlike 'APPCATALOG*' -and
                $_.Template -notlike 'REDIRECTSITE*' -and
                $_.Url -notlike '*-my.sharepoint.com/personal/*'
            }
            
            $siteUrls = @($allSites | Select-Object -ExpandProperty Url)
            Write-Host "Found $($siteUrls.Count) SharePoint sites" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Found $($siteUrls.Count) SharePoint sites" -LogLevel "INFO"
        }
        elseif ($Scope -eq "2") {
            # Get only OneDrive for Business sites
            Write-Host "Retrieving OneDrive for Business sites..." -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Retrieving OneDrive sites" -LogLevel "INFO"
            
            $allSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"
            
            $siteUrls = @($allSites | Select-Object -ExpandProperty Url)
            Write-Host "Found $($siteUrls.Count) OneDrive sites" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Found $($siteUrls.Count) OneDrive sites" -LogLevel "INFO"
        }
        
        return $siteUrls
    }
    catch {
        $errorMsg = "Failed to discover sites: $_"
        Write-Error $errorMsg
        Write-Host $_.Exception.ToString() -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        return $null
    }
}

# Function to filter sites for version report-related operations by current site size
function Get-VersionReportEligibleSites {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$SiteUrls,

        [Parameter(Mandatory = $false)]
        [double]$MinSiteSizeMB = 0
    )

    if ($MinSiteSizeMB -le 0) {
        Write-Host "Version report site-size filter disabled (MinSiteSizeforversionReports <= 0)." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "Version report site-size filter disabled. Processing all $($SiteUrls.Count) sites." -LogLevel "INFO"
        return [PSCustomObject]@{
            EligibleSiteUrls = @($SiteUrls)
            SkippedSites     = @()
            UnknownSites     = @()
        }
    }

    Write-Host "`nApplying version report site-size filter (StorageUsageCurrent >= $MinSiteSizeMB MB)..." -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Applying version report site-size filter: StorageUsageCurrent >= $MinSiteSizeMB MB" -LogLevel "INFO"

    $allSites = Get-PnPTenantSite -IncludeOneDriveSites
    $siteByUrl = @{}
    foreach ($tenantSite in $allSites) {
        $normalizedUrl = $tenantSite.Url.TrimEnd('/').ToLowerInvariant()
        $siteByUrl[$normalizedUrl] = $tenantSite
    }

    $eligible = [System.Collections.Generic.List[string]]::new()
    $skipped  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $unknown  = [System.Collections.Generic.List[string]]::new()

    foreach ($siteUrl in $SiteUrls) {
        $cleanUrl = $siteUrl.TrimEnd('/')
        $normalizedUrl = $cleanUrl.ToLowerInvariant()

        if ($siteByUrl.ContainsKey($normalizedUrl)) {
            $siteInfo = $siteByUrl[$normalizedUrl]
            $storageUsageMB = [double]$siteInfo.StorageUsageCurrent
            if ($storageUsageMB -ge $MinSiteSizeMB) {
                $eligible.Add($cleanUrl)
            }
            else {
                $skipped.Add([PSCustomObject]@{
                    SiteUrl          = $cleanUrl
                    StorageUsageMB   = [math]::Round($storageUsageMB, 2)
                    MinimumSizeMB    = $MinSiteSizeMB
                })
            }
        }
        else {
            # Keep unknown sites to avoid accidentally excluding valid targets when metadata lookup misses.
            $eligible.Add($cleanUrl)
            $unknown.Add($cleanUrl)
        }
    }

    Write-Host "Eligible sites for version reports: $($eligible.Count)" -ForegroundColor Green
    Write-Host "Skipped by size filter         : $($skipped.Count)" -ForegroundColor Yellow
    if ($unknown.Count -gt 0) {
        Write-Host "Included without size metadata : $($unknown.Count)" -ForegroundColor Yellow
    }

    Write-LogEntry -LogName $log -LogEntryText "Version report size filter results: Eligible=$($eligible.Count), Skipped=$($skipped.Count), UnknownIncluded=$($unknown.Count), ThresholdMB=$MinSiteSizeMB" -LogLevel "INFO"

    return [PSCustomObject]@{
        EligibleSiteUrls = @($eligible)
        SkippedSites     = @($skipped)
        UnknownSites     = @($unknown)
    }
}

# Log script start
Write-LogEntry -LogName $log -LogEntryText "Script execution started. Connecting to tenant admin site: $url" -LogLevel "INFO"

# Connect to the SharePoint Online admin site
try {
    Connect-PnPOnline -Url $url -ClientId $clientId -Tenant $tenantId -Interactive
    $connection = Get-PnPConnection
    Write-LogEntry -LogName $log -LogEntryText "Successfully connected to admin site" -LogLevel "INFO"
}
catch {
    Write-Host "ERROR: Failed to connect to the SharePoint admin center." -ForegroundColor Red
    Write-Host "  URL    : $url" -ForegroundColor Red
    Write-Host "  Detail : $_" -ForegroundColor Red
    Write-LogEntry -LogName $log -LogEntryText "Failed to connect to admin site: $_" -LogLevel "ERROR"
    exit
}

# Load or discover sites based on configuration
$sites = $null

if ($null -ne $sitesFilePath -and $sitesFilePath -ne "") {
    # Load sites from file
    if (Test-Path $sitesFilePath) {
        # Filter blank and whitespace-only lines so stray newlines in the file don't create empty site URLs
        $sites = @(Get-Content -Path $sitesFilePath | Where-Object { $_.Trim() -ne '' })
        if ($sites.Count -eq 0) {
            Write-Host "WARNING: Site list file exists but contains no valid URLs: $sitesFilePath" -ForegroundColor Yellow
            Write-LogEntry -LogName $log -LogEntryText "Site list file is empty: $sitesFilePath" -LogLevel "WARNING"
            Write-Host "`nExiting script..." -ForegroundColor Red
            exit
        }
        Write-Host "Loaded $($sites.Count) sites from file: $sitesFilePath" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Reading site list from: $sitesFilePath" -LogLevel "INFO"
        Write-LogEntry -LogName $log -LogEntryText "Found $($sites.Count) sites to process" -LogLevel "INFO"
    }
    else {
        Write-Host "WARNING: Site list file not found at: $sitesFilePath" -ForegroundColor Yellow
        Write-Host "The file path is configured but the file does not exist." -ForegroundColor Yellow
        Write-Host "Please either:" -ForegroundColor Yellow
        Write-Host "  1. Create the file with site URLs (one per line), or" -ForegroundColor Yellow
        Write-Host "  2. Set `$sitesFilePath = `$null in the script to auto-discover sites" -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "Site list file not found: $sitesFilePath" -LogLevel "ERROR"
        Write-Host "`nExiting script..." -ForegroundColor Red
        exit
    }
}
else {
    # Auto-discovery mode
    Write-Host "`n==== Site Discovery Mode ====" -ForegroundColor Cyan
    Write-Host "No site list file configured. The script will discover sites automatically." -ForegroundColor Yellow
    Write-Host "This is recommended for smaller tenants." -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Site auto-discovery mode enabled (sitesFilePath is null)" -LogLevel "INFO"
    
    # This will be populated when user selects an operation
    Write-Host "You will be prompted to select SharePoint sites or OneDrive sites before each operation." -ForegroundColor Cyan
}



# Function to handle throttling
function Invoke-WithThrottlingHandling {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [int]$MaxRetries = 5,
        [int]$InitialRetrySeconds = 30
    )
    
    $retryCount = 0
    $success = $false
    
    Write-Host "Executing operation on site: $SiteUrl" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Executing operation on site: $SiteUrl" -LogLevel "INFO"
    
    while (-not $success -and $retryCount -lt $MaxRetries) {
        try {
            # Execute the command and capture output
            $output = & $ScriptBlock
            $success = $true
            
            # Display the output to console if it's not empty
            if ($output) {
                Write-Host "Output from site $SiteUrl :" -ForegroundColor Green
                $output | Format-Table -AutoSize
            }
            
            Write-Host "Successfully executed command for site: $SiteUrl" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Successfully executed command for site: $SiteUrl" -LogLevel "INFO"
        }
        catch {
            # Guard against exceptions that don't expose .Response (e.g. CSOM, network errors)
            $statusCode = $null
            try { $statusCode = $_.Exception.Response.StatusCode } catch { }
            if ($statusCode -eq 429 -or $statusCode -eq 503) {
                $retryAfter = $null
                try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
                if (-not $retryAfter) {
                    $retryAfter = $InitialRetrySeconds * [math]::Pow(2, $retryCount)
                }
                
                $retryCount++
                $warningMsg = "Throttling detected for site $SiteUrl. Waiting for $retryAfter seconds before retry $retryCount of $MaxRetries..."
                Write-Warning $warningMsg
                Write-LogEntry -LogName $log -LogEntryText $warningMsg -LogLevel "WARNING"
                Start-Sleep -Seconds ([int][math]::Ceiling($retryAfter))
            }
            else {
                $errorMsg = "Error processing site $SiteUrl : $_"
                Write-Error $errorMsg
                Write-Host $_.Exception.ToString() -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                throw $_
            }
        }
    }
    
    if (-not $success) {
        $errorMsg = "Failed to execute command for $SiteUrl after $MaxRetries retries."
        Write-Error $errorMsg
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
    }
}

# Lightweight retry wrapper that returns the scriptblock result; used by tenant-level and per-site loop functions
function Invoke-PnPWithRetry {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [string]$OperationDescription = "operation",
        [int]$MaxRetries = 5,
        [int]$InitialRetrySeconds = 30
    )

    $retryCount = 0
    while ($true) {
        try {
            return (& $ScriptBlock)
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
            if (($statusCode -eq 429 -or $statusCode -eq 503) -and $retryCount -lt $MaxRetries) {
                $retryAfter = $null
                try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
                if (-not $retryAfter) {
                    $retryAfter = $InitialRetrySeconds * [math]::Pow(2, $retryCount)
                }
                $retryCount++
                $msg = "Throttling detected ($OperationDescription). Waiting $([int][math]::Ceiling($retryAfter))s before retry $retryCount/$MaxRetries..."
                Write-Warning $msg
                Write-LogEntry -LogName $log -LogEntryText $msg -LogLevel "WARNING"
                Start-Sleep -Seconds ([int][math]::Ceiling($retryAfter))
            }
            else {
                throw $_
            }
        }
    }
}

# Function to process each site with a specific operation
function Invoke-SiteBatch {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$SiteUrls,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [object]$Connection,
        
        [string]$OperationDescription = "operation"
    )
    
    # If SiteUrls is null or empty, prompt for site discovery
    if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
        Write-Host "`nNo sites loaded. Starting site discovery..." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "Site discovery triggered for operation: $OperationDescription" -LogLevel "INFO"
        
        $scopeChoice = Get-SiteScope
        
        if ($scopeChoice -eq "3") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            Write-LogEntry -LogName $log -LogEntryText "User cancelled site discovery" -LogLevel "INFO"
            return
        }
        
        $SiteUrls = Get-FilteredSites -Scope $scopeChoice
        
        if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
            Write-Host "No sites found or discovery failed. Operation cancelled." -ForegroundColor Red
            Write-LogEntry -LogName $log -LogEntryText "No sites discovered for operation: $OperationDescription" -LogLevel "WARNING"
            return
        }
        
        # Confirm with user before proceeding
        Write-Host "`nReady to process $($SiteUrls.Count) sites." -ForegroundColor Yellow
        $confirm = Read-Host "Proceed with operation? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            Write-LogEntry -LogName $log -LogEntryText "User cancelled operation after site discovery" -LogLevel "INFO"
            return
        }
    }
    
    Write-Host "Starting batch processing for operation: $OperationDescription on $($SiteUrls.Count) sites" -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Starting batch processing for operation: $OperationDescription on $($SiteUrls.Count) sites" -LogLevel "INFO"
    
    foreach ($siteUrl in $SiteUrls) {
        Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "Processing site: $siteUrl for $OperationDescription" -LogLevel "INFO"
        
        try {
            # Connect to the site using the existing authentication (no interactive prompt)
            Write-Host "Connecting to site: $siteUrl" -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Connecting to site: $siteUrl" -LogLevel "DEBUG"
            Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -Interactive
            
            # Apply site operation with throttling handling
            Invoke-WithThrottlingHandling -SiteUrl $siteUrl -ScriptBlock $Operation
        }
        catch {
            $errorMsg = "Failed to connect to site $siteUrl. Error: $_"
            Write-Error $errorMsg
            Write-Host $_.Exception.ToString() -ForegroundColor Red
            Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        }
    }
    
    Write-Host "Processing completed for all sites" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Completed batch processing for operation: $OperationDescription" -LogLevel "INFO"
}

# Create operation script blocks
$getVersionPolicyOperation = {
    $policy = Get-PnPSiteVersionPolicy
    # Return policy object for display
    Write-Host "  - Site version policy retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy retrieved: EnableAutoExpirationVersionTrim = $($policy.DefaultTrimMode)" -LogLevel "INFO"
    return $policy | Format-List # Format list for better readability
}

$setVersionPolicyOperation = {
    $result = Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $true
    Write-Host "  - Site version policy set successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy set to EnableAutoExpirationVersionTrim = True" -LogLevel "INFO"
    return $result | Format-List # Format list for better readability
}

$getVersionPolicyStatusOperation = {
    $status = Get-PnPSiteVersionPolicyStatus
    # Return status object for display
    Write-Host "  - Site version policy status retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Site version policy status:  $($status.Status), CompleteTimeInUTC:  $($status.CompleteTimeInUTC)" -LogLevel "INFO"
    return $status | Format-List
}

$createBatchDeleteJobOperation = {
    $job = New-PnPSiteFileVersionBatchDeleteJob -Automatic -Force
    # Return job object for display
    Write-Host "Site file version batch delete job created successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Batch delete job created with status $($job)" -LogLevel "INFO"
    return $job | Format-List # Format list for better readability
}

# Create manual batch delete job operation (will be populated with user settings)
$createManualBatchDeleteJobOperation = $null

$getBatchDeleteJobStatusOperation = {
    $jobStatus = Get-PnPSiteFileVersionBatchDeleteJobStatus
    # Return job status object for display
    Write-Host "  - Site file version batch delete job status retrieved successfully" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Batch delete job status: State = $($jobStatus.Status), CompleteTimeInUTC = $($jobStatus.CompleteTimeInUTC), BatchDeleteMode = $($jobStatus.BatchDeleteMode), StorageReleasedInBytes = $($jobStatus.StorageReleasedInBytes)"  -LogLevel "INFO"
    return $jobStatus | Format-List # Format list for better readability
}

# Function to prompt for batch delete settings (manual mode)
function Get-BatchDeleteSettings {
    Write-Host "`n==== Configure Batch Delete Settings ====" -ForegroundColor Cyan
    Write-Host "Manual deletion allows you to specify version count or age limits." -ForegroundColor Cyan
    
    # Get deletion mode preference
    Write-Host "`n==== Select Deletion Mode ====" -ForegroundColor Cyan
    Write-Host "1: Delete by age (DeleteOlderThanDays) - removes all versions older than specified days"
    Write-Host "2: Delete by count (MajorVersionLimit) - keeps only specified number of most recent versions"
    
    $modeChoice = $null
    do {
        $modeChoice = Read-Host "Select deletion mode (1-2)"
        if ($modeChoice -notin @("1", "2")) {
            Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
        }
    } while ($modeChoice -notin @("1", "2"))
    
    $deleteSettings = @{}
    
    if ($modeChoice -eq "1") {
        # Delete by age
        Write-Host "`n==== Delete Versions Older Than ====" -ForegroundColor Cyan
        Write-Host "1: 30 days"
        Write-Host "2: 90 days"
        Write-Host "3: 180 days"
        Write-Host "4: 365 days"
        Write-Host "5: Custom (must be at least 30 days)"
        
        do {
            $ageChoice = Read-Host "Select age (1-5)"
            $validChoice = $true
            $deleteOlderThanDays = 0
            
            switch ($ageChoice) {
                "1" {
                    $deleteOlderThanDays = 30
                    Write-Host "Selected: 30 days" -ForegroundColor Green
                }
                "2" {
                    $deleteOlderThanDays = 90
                    Write-Host "Selected: 90 days" -ForegroundColor Green
                }
                "3" {
                    $deleteOlderThanDays = 180
                    Write-Host "Selected: 180 days" -ForegroundColor Green
                }
                "4" {
                    $deleteOlderThanDays = 365
                    Write-Host "Selected: 365 days" -ForegroundColor Green
                }
                "5" {
                    do {
                        $customDaysInput = Read-Host "Enter number of days (minimum 30)"
                        $customDays = $null
                        $validCustomDays = [int]::TryParse($customDaysInput, [ref]$customDays)
                        
                        if (-not $validCustomDays -or $customDays -lt 30) {
                            Write-Host "Invalid input. Please enter a number of 30 or greater." -ForegroundColor Red
                        }
                        else {
                            $deleteOlderThanDays = $customDays
                            Write-Host "Selected: Custom ($customDays days)" -ForegroundColor Green
                        }
                    } while (-not $validCustomDays -or $customDays -lt 30)
                }
                default {
                    Write-Host "Invalid selection. Please choose 1-5." -ForegroundColor Red
                    $validChoice = $false
                }
            }
        } while (-not $validChoice)
        
        $deleteSettings.DeleteOlderThanDays = $deleteOlderThanDays
        Write-LogEntry -LogName $log -LogEntryText "User set batch delete by age: DeleteOlderThanDays = $deleteOlderThanDays" -LogLevel "INFO"
    }
    else {
        # Delete by major version limit - keeping X most recent versions
        Write-Host "`n==== Specify Version Count Limit ====" -ForegroundColor Cyan
        Write-Host "Enter how many recent major versions to KEEP (older versions will be deleted)" -ForegroundColor Yellow
        
        do {
            $majorVersionInput = Read-Host "Enter the major version limit to keep (minimum 100)"
            $majorVersionLimit = $null
            $validMajorVersion = [int]::TryParse($majorVersionInput, [ref]$majorVersionLimit)
            
            if (-not $validMajorVersion -or $majorVersionLimit -lt 100) {
                Write-Host "Invalid input. Please enter a positive integer of 100 or greater." -ForegroundColor Red
            }
        } while (-not $validMajorVersion -or $majorVersionLimit -lt 100)
        
        $deleteSettings.MajorVersionLimit = $majorVersionLimit
        Write-Host "Will keep $majorVersionLimit most recent versions and delete older ones" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "User set batch delete by version count: MajorVersionLimit = $majorVersionLimit" -LogLevel "INFO"
    }
    
    return $deleteSettings
}

# Function to prompt for manual version settings
function Get-ManualVersionSettings {
    Write-Host "`n==== Configure Manual Version Settings ====" -ForegroundColor Cyan
    
    # Get major version limit
    do {
        $majorVersionInput = Read-Host "Enter the major version limit (minimum 100)"
        $majorVersionLimit = $null
        $validMajorVersion = [int]::TryParse($majorVersionInput, [ref]$majorVersionLimit)
        
        if (-not $validMajorVersion -or $majorVersionLimit -lt 100) {
            Write-Host "Invalid input. Please enter a positive integer of 100 or greater." -ForegroundColor Red
        }
    } while (-not $validMajorVersion -or $majorVersionLimit -lt 100)
    
    Write-LogEntry -LogName $log -LogEntryText "User set major version limit: $majorVersionLimit" -LogLevel "INFO"
    
    # Get time setting
    Write-Host "`n==== Select Time Setting ====" -ForegroundColor Cyan
    Write-Host "1: Never (Default)"
    Write-Host "2: 3 months (90 days)"
    Write-Host "3: 6 months (180 days)"
    Write-Host "4: 1 year (365 days)"
    Write-Host "5: Custom (must be greater than 29 days)"
    
    do {
        $timeChoice = Read-Host "Select time setting (1-5)"
        $validChoice = $true
        $expireAfterDays = $null
        
        switch ($timeChoice) {
            "1" {
                $expireAfterDays = $null
                Write-Host "Selected: Never (Default)" -ForegroundColor Green
            }
            "2" {
                $expireAfterDays = 90
                Write-Host "Selected: 3 months (90 days)" -ForegroundColor Green
            }
            "3" {
                $expireAfterDays = 180
                Write-Host "Selected: 6 months (180 days)" -ForegroundColor Green
            }
            "4" {
                $expireAfterDays = 365
                Write-Host "Selected: 1 year (365 days)" -ForegroundColor Green
            }
            "5" {
                do {
                    $customDaysInput = Read-Host "Enter custom number of days (must be greater than 29)"
                    $customDays = $null
                    $validCustomDays = [int]::TryParse($customDaysInput, [ref]$customDays)
                    
                    if (-not $validCustomDays -or $customDays -le 29) {
                        Write-Host "Invalid input. Please enter a number greater than 29." -ForegroundColor Red
                    }
                    else {
                        $expireAfterDays = $customDays
                        Write-Host "Selected: Custom ($customDays days)" -ForegroundColor Green
                    }
                } while (-not $validCustomDays -or $customDays -le 29)
            }
            default {
                Write-Host "Invalid selection. Please choose 1-5." -ForegroundColor Red
                $validChoice = $false
            }
        }
    } while (-not $validChoice)
    
    Write-LogEntry -LogName $log -LogEntryText "User set time setting: $(if ($expireAfterDays) { "$expireAfterDays days" } else { "Never (Default)" })" -LogLevel "INFO"
    
    # Return settings as hashtable
    return @{
        MajorVersionLimit = $majorVersionLimit
        ExpireAfterDays   = $expireAfterDays
    }
}

# Create manual version policy operation (will be populated with user settings)
$setManualVersionPolicyOperation = $null

# Function to set tenant-level automatic version settings
function Set-TenantAutomaticVersionPolicy {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AdminUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    Write-Host "`n==== Setting Tenant-Level Automatic Version Policy ====" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting tenant-level automatic version policy configuration" -LogLevel "INFO"
    
    try {
        # Ensure we're connected to the admin site
        if ($AdminUrl -and $ClientId -and $TenantId) {
            try {
                $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
                $needsReconnect = $true
                
                if ($currentConnection) {
                    # Check if we're connected to the admin URL
                    if ($currentConnection.Url -eq $AdminUrl) {
                        $needsReconnect = $false
                    }
                }
                
                if ($needsReconnect) {
                    Write-LogEntry -LogName $log -LogEntryText "Reconnecting to admin site: $AdminUrl" -LogLevel "DEBUG"
                    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
                }
            }
            catch {
                Write-LogEntry -LogName $log -LogEntryText "Connection check failed, reconnecting: $_" -LogLevel "DEBUG"
                Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
            }
        }
        
        # Set tenant to automatic mode
        Invoke-PnPWithRetry -OperationDescription "set tenant automatic version policy" -ScriptBlock {
            Set-PnPTenant -EnableAutoExpirationVersionTrim $true
        }
        
        Write-Host "Successfully set tenant to Automatic version trimming mode" -ForegroundColor Green
        Write-Host "New sites will automatically optimize storage using an intelligent algorithm." -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Tenant-level automatic version policy set: EnableAutoExpirationVersionTrim = True" -LogLevel "INFO"
        
        return $true
    }
    catch {
        $errorMsg = "Failed to set tenant-level automatic version policy: $_"
        Write-Error $errorMsg
        Write-Host $_.Exception.ToString() -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        return $false
    }
}

# Function to review current tenant-level version settings
function Get-TenantVersionSettings {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AdminUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    Write-Host "`n==== Retrieving Tenant-Level Version Settings ====" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Retrieving tenant-level version settings for review" -LogLevel "INFO"
    
    try {
        # Ensure we're connected to the admin site
        if ($AdminUrl -and $ClientId -and $TenantId) {
            try {
                $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
                $needsReconnect = $true
                
                if ($currentConnection) {
                    # Check if we're connected to the admin URL
                    if ($currentConnection.Url -eq $AdminUrl) {
                        $needsReconnect = $false
                    }
                }
                
                if ($needsReconnect) {
                    Write-LogEntry -LogName $log -LogEntryText "Reconnecting to admin site: $AdminUrl" -LogLevel "DEBUG"
                    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
                }
            }
            catch {
                Write-LogEntry -LogName $log -LogEntryText "Connection check failed, reconnecting: $_" -LogLevel "DEBUG"
                Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
            }
        }
        
        $tenantConfig = Invoke-PnPWithRetry -OperationDescription "get tenant version settings" -ScriptBlock {
            Get-PnPTenant
        }
        
        Write-Host "`n==== Current Tenant Version Settings ====" -ForegroundColor Cyan
        Write-Host ""
        
        # Display version policy mode
        if ($tenantConfig.EnableAutoExpirationVersionTrim -eq $true) {
            Write-Host "Version Policy Mode:" -ForegroundColor Yellow
            Write-Host "  Automatic Version Trimming: ENABLED" -ForegroundColor Green
            Write-Host "  Description: Uses intelligent algorithm to optimize storage based on version creation date" -ForegroundColor White
        }
        else {
            Write-Host "Version Policy Mode:" -ForegroundColor Yellow
            Write-Host "  Manual Version Limits: ENABLED" -ForegroundColor Green
            Write-Host ""
            Write-Host "Manual Version Settings:" -ForegroundColor Yellow
            Write-Host "  Major Version Limit: $($tenantConfig.MajorVersionLimit)" -ForegroundColor White
            
            if ($tenantConfig.ExpireVersionsAfterDays -eq 0) {
                Write-Host "  Expire After Days: Never (No Expiration)" -ForegroundColor White
            }
            else {
                Write-Host "  Expire After Days: $($tenantConfig.ExpireVersionsAfterDays) days" -ForegroundColor White
            }
        }
        
        Write-Host ""
        Write-Host "What this means:" -ForegroundColor Yellow
        Write-Host "  - New sites created in this tenant will inherit these settings" -ForegroundColor White
        Write-Host "  - Existing sites retain their individual settings unless explicitly changed" -ForegroundColor White
        Write-Host "  - Use Option 2 to apply these settings to existing sites" -ForegroundColor White
        
        Write-LogEntry -LogName $log -LogEntryText "Displayed tenant settings: EnableAutoExpirationVersionTrim = $($tenantConfig.EnableAutoExpirationVersionTrim), MajorVersionLimit = $($tenantConfig.MajorVersionLimit), ExpireVersionsAfterDays = $($tenantConfig.ExpireVersionsAfterDays)" -LogLevel "INFO"
        
        return $true
    }
    catch {
        $errorMsg = "Failed to retrieve tenant settings: $_"
        Write-Error $errorMsg
        Write-Host $_.Exception.ToString() -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        return $false
    }
}

# Function to set tenant-level manual version settings
function Set-TenantManualVersionPolicy {
    param (
        [Parameter(Mandatory = $false)]
        [string]$AdminUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    Write-Host "`n==== Setting Tenant-Level Manual Version Policy ====" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting tenant-level manual version policy configuration" -LogLevel "INFO"
    
    # Get user input for manual version settings
    $tenantSettings = Get-ManualVersionSettings
    
    try {
        # Ensure we're connected to the admin site
        if ($AdminUrl -and $ClientId -and $TenantId) {
            try {
                $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
                $needsReconnect = $true
                
                if ($currentConnection) {
                    # Check if we're connected to the admin URL
                    if ($currentConnection.Url -eq $AdminUrl) {
                        $needsReconnect = $false
                    }
                }
                
                if ($needsReconnect) {
                    Write-LogEntry -LogName $log -LogEntryText "Reconnecting to admin site: $AdminUrl" -LogLevel "DEBUG"
                    Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
                }
            }
            catch {
                Write-LogEntry -LogName $log -LogEntryText "Connection check failed, reconnecting: $_" -LogLevel "DEBUG"
                Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Tenant $TenantId -Interactive | Out-Null
            }
        }
        
        # Build parameters for Set-PnPTenant
        $params = @{
            EnableAutoExpirationVersionTrim = $false
            MajorVersionLimit               = $tenantSettings.MajorVersionLimit
        }
        
        # Handle ExpireVersionsAfterDays
        # If null (Never), set to 0 for NoExpiration
        # Otherwise use the value (must be >= 30 for ExpireAfter according to API)
        if ($null -eq $tenantSettings.ExpireAfterDays) {
            $params.ExpireVersionsAfterDays = 0
            $expireDisplay = "Never (No Expiration)"
        }
        else {
            # Ensure it's at least 30 days for the tenant setting
            if ($tenantSettings.ExpireAfterDays -lt 30) {
                Write-Host "Note: Tenant-level setting requires minimum 30 days. Adjusting from $($tenantSettings.ExpireAfterDays) to 30 days." -ForegroundColor Yellow
                $params.ExpireVersionsAfterDays = 30
                $expireDisplay = "30 days (minimum for tenant setting)"
            }
            else {
                $params.ExpireVersionsAfterDays = $tenantSettings.ExpireAfterDays
                $expireDisplay = "$($tenantSettings.ExpireAfterDays) days"
            }
        }
        
        # Set tenant manual version policy
        Invoke-PnPWithRetry -OperationDescription "set tenant manual version policy" -ScriptBlock ({
            Set-PnPTenant @params
        }.GetNewClosure())
        
        Write-Host "`nSuccessfully set tenant to Manual version limits mode" -ForegroundColor Green
        Write-Host "  Major Version Limit: $($tenantSettings.MajorVersionLimit)" -ForegroundColor Green
        Write-Host "  Expire After Days: $expireDisplay" -ForegroundColor Green
        Write-Host "`nNew sites will use these version limits by default." -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "Tenant-level manual version policy set: EnableAutoExpirationVersionTrim = False, MajorVersionLimit = $($tenantSettings.MajorVersionLimit), ExpireVersionsAfterDays = $($params.ExpireVersionsAfterDays)" -LogLevel "INFO"
        
        return $true
    }
    catch {
        $errorMsg = "Failed to set tenant-level manual version policy: $_"
        Write-Error $errorMsg
        Write-Host $_.Exception.ToString() -ForegroundColor Red
        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
        return $false
    }
}

# Function to generate version expiration reports for all site collections
function New-TenantVersionExpirationReport {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$SiteUrls,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )
    
    Write-Host "`n==== Generate Version History Report for All Site Collections ====" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting version expiration report generation for all sites" -LogLevel "INFO"
    
    Write-Host "`nReports will be saved to the 'Admin_SiteCollection_VersionReport_DONOTDELETE' library in each site collection." -ForegroundColor Yellow
    Write-Host "Filename: {SiteCollectionName}site_adminreport_donotdelete_VersionReport.csv" -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Report destination: Admin_SiteCollection_VersionReport_DONOTDELETE library in each site collection" -LogLevel "INFO"
    
    # Constant for the dedicated report library name
    $reportLibraryName = "Admin_SiteCollection_VersionReport_DONOTDELETE"
    
    # Resolve site list if not provided (auto-discovery mode)
    if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
        Write-Host "`nNo sites loaded. Starting site discovery..." -ForegroundColor Yellow
        $scopeChoice = Get-SiteScope
        if ($scopeChoice -eq "3") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return
        }
        $SiteUrls = Get-FilteredSites -Scope $scopeChoice
        if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
            Write-Host "No sites found or discovery failed. Operation cancelled." -ForegroundColor Red
            return
        }
        $confirm = Read-Host "`nReady to process $($SiteUrls.Count) sites. Proceed? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return
        }
    }

    $filteredSites = Get-VersionReportEligibleSites -SiteUrls $SiteUrls -MinSiteSizeMB $script:MinSiteSizeforversionReports
    $SiteUrls = $filteredSites.EligibleSiteUrls

    if ($SiteUrls.Count -eq 0) {
        Write-Host "No sites meet the minimum size threshold ($($script:MinSiteSizeforversionReports) MB)." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "No sites eligible for report generation after size filter. ThresholdMB=$($script:MinSiteSizeforversionReports)" -LogLevel "WARNING"
        return
    }
    
    Write-Host "`nStarting report generation for $($SiteUrls.Count) sites..." -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Starting report generation for $($SiteUrls.Count) sites" -LogLevel "INFO"
    
    foreach ($siteUrl in $SiteUrls) {
        $cleanUrl = $siteUrl.TrimEnd('/')
        Write-Host "`nProcessing site: $cleanUrl" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "Processing report for site: $cleanUrl" -LogLevel "INFO"
        
        $siteRetry = 0
        $siteDone = $false
        while (-not $siteDone) {
            try {
                $siteCollectionName = ($cleanUrl -split '/') | Where-Object { $_ -ne '' } | Select-Object -Last 1

                Connect-PnPOnline -Url $cleanUrl -ClientId $ClientId -Tenant $TenantId -Interactive

                # Create the dedicated report library if it doesn't already exist
                $reportLib = Get-PnPList -Identity $reportLibraryName -ErrorAction SilentlyContinue
                if ($null -eq $reportLib) {
                    Write-Host "  - Creating report library: $reportLibraryName" -ForegroundColor Cyan
                    New-PnPList -Title $reportLibraryName -Template DocumentLibrary -ErrorAction Stop | Out-Null
                    Write-LogEntry -LogName $log -LogEntryText "Created report library '$reportLibraryName' on $cleanUrl" -LogLevel "INFO"
                }
                else {
                    Write-Host "  - Report library already exists: $reportLibraryName" -ForegroundColor Cyan
                }

                $reportFileName = "${siteCollectionName}site_adminreport_donotdelete_VersionReport.csv"
                $fullReportUrl  = "$cleanUrl/$reportLibraryName/$reportFileName"

                # Delete existing report file so the job can be resubmitted without error
                $existingReport = Get-PnPFile -Url "/$reportLibraryName/$reportFileName" -ErrorAction SilentlyContinue
                if ($null -ne $existingReport) {
                    Write-Host "  - Existing report found — deleting before resubmitting" -ForegroundColor Yellow
                    Write-LogEntry -LogName $log -LogEntryText "Deleting existing report file before resubmit: $fullReportUrl" -LogLevel "INFO"
                    Remove-PnPFile -SiteRelativeUrl "/$reportLibraryName/$reportFileName" -Force -ErrorAction Stop
                }

                Write-Host "  - Submitting version expiration report job" -ForegroundColor Cyan
                Write-Host "    Report URL: $fullReportUrl" -ForegroundColor Cyan
                Write-LogEntry -LogName $log -LogEntryText "Submitting report job. ReportUrl: $fullReportUrl" -LogLevel "INFO"

                New-PnPSiteFileVersionExpirationReportJob -ReportUrl $fullReportUrl

                Write-Host "  - Report job submitted successfully: $reportFileName" -ForegroundColor Green
                Write-LogEntry -LogName $log -LogEntryText "Report job submitted for site: $cleanUrl, Filename: $reportFileName" -LogLevel "INFO"
                $siteDone = $true
            }
            catch {
                $statusCode = $null
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                if (($statusCode -eq 429 -or $statusCode -eq 503) -and $siteRetry -lt 5) {
                    $retryAfter = $null
                    try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
                    if (-not $retryAfter) { $retryAfter = 30 * [math]::Pow(2, $siteRetry) }
                    $siteRetry++
                    $msg = "Throttling detected for $cleanUrl. Waiting $([int][math]::Ceiling($retryAfter))s before retry $siteRetry/5..."
                    Write-Warning $msg
                    Write-LogEntry -LogName $log -LogEntryText $msg -LogLevel "WARNING"
                    Start-Sleep -Seconds ([int][math]::Ceiling($retryAfter))
                }
                else {
                    $errorMsg = "Failed to submit report job for $cleanUrl : $_"
                    Write-Error $errorMsg
                    Write-Host $_.Exception.ToString() -ForegroundColor Red
                    Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                    $siteDone = $true
                }
            }
        }
    }
    
    Write-Host "`nReport generation completed for all sites." -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "Completed report generation for all sites" -LogLevel "INFO"
}

# Function to check version expiration report job status for all site collections
function Get-TenantVersionExpirationReportStatus {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$SiteUrls,

        [Parameter(Mandatory = $false)]
        [string]$ClientId,

        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )

    Write-Host "`n==== Get Version History Report Job Status for All Site Collections ====" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Starting version expiration report status check for all sites" -LogLevel "INFO"

    # Resolve site list if not provided (auto-discovery mode)
    if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
        Write-Host "`nNo sites loaded. Starting site discovery..." -ForegroundColor Yellow
        $scopeChoice = Get-SiteScope
        if ($scopeChoice -eq "3") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return
        }
        $SiteUrls = Get-FilteredSites -Scope $scopeChoice
        if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
            Write-Host "No sites found or discovery failed. Operation cancelled." -ForegroundColor Red
            return
        }
        $confirm = Read-Host "`nReady to process $($SiteUrls.Count) sites. Proceed? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            return
        }
    }

    $filteredSites = Get-VersionReportEligibleSites -SiteUrls $SiteUrls -MinSiteSizeMB $script:MinSiteSizeforversionReports
    $SiteUrls = $filteredSites.EligibleSiteUrls
    if ($SiteUrls.Count -eq 0) {
        Write-Host "No sites meet the minimum size threshold ($($script:MinSiteSizeforversionReports) MB)." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "No sites eligible for report status check after size filter. ThresholdMB=$($script:MinSiteSizeforversionReports)" -LogLevel "WARNING"
        return
    }

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($siteUrl in $SiteUrls) {
        $cleanUrl = $siteUrl.TrimEnd('/')
        Write-Host "`nProcessing site: $cleanUrl" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "Checking report status for site: $cleanUrl" -LogLevel "INFO"

        $siteRetry = 0
        $siteDone = $false
        while (-not $siteDone) {
            try {
                $siteCollectionName = ($cleanUrl -split '/') | Where-Object { $_ -ne '' } | Select-Object -Last 1
                $reportFileName     = "${siteCollectionName}site_adminreport_donotdelete_VersionReport.csv"
                $reportLibraryName  = "Admin_SiteCollection_VersionReport_DONOTDELETE"
                $fullReportUrl      = "$cleanUrl/$reportLibraryName/$reportFileName"

                Connect-PnPOnline -Url $cleanUrl -ClientId $ClientId -Tenant $TenantId -Interactive

                Write-Host "  - Checking report job status" -ForegroundColor Cyan
                Write-Host "    Report URL: $fullReportUrl" -ForegroundColor Cyan

                $status = Get-PnPSiteFileVersionExpirationReportJobStatus -ReportUrl $fullReportUrl

                $statusValue = $status.Status
                $errorMsg    = $status.ErrorMessage

                switch ($statusValue) {
                    "completed"  { Write-Host "  - Status: Completed" -ForegroundColor Green }
                    "failed"     { Write-Host "  - Status: Failed - $errorMsg" -ForegroundColor Red }
                    default      { Write-Host "  - Status: $statusValue" -ForegroundColor Yellow }
                }

                $results.Add([PSCustomObject]@{
                    SiteUrl      = $cleanUrl
                    Status       = $statusValue
                    ErrorMessage = $errorMsg
                })

                Write-LogEntry -LogName $log -LogEntryText "Report status for $cleanUrl : $statusValue $(if ($errorMsg) { "- $errorMsg" })" -LogLevel "INFO"
                $siteDone = $true
            }
            catch {
                $statusCode = $null
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                if (($statusCode -eq 429 -or $statusCode -eq 503) -and $siteRetry -lt 5) {
                    $retryAfter = $null
                    try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
                    if (-not $retryAfter) { $retryAfter = 30 * [math]::Pow(2, $siteRetry) }
                    $siteRetry++
                    $msg = "Throttling detected for $cleanUrl. Waiting $([int][math]::Ceiling($retryAfter))s before retry $siteRetry/5..."
                    Write-Warning $msg
                    Write-LogEntry -LogName $log -LogEntryText $msg -LogLevel "WARNING"
                    Start-Sleep -Seconds ([int][math]::Ceiling($retryAfter))
                }
                else {
                    $errText = "Failed to get report status for $cleanUrl : $_"
                    Write-Error $errText
                    Write-Host $_.Exception.ToString() -ForegroundColor Red
                    Write-LogEntry -LogName $log -LogEntryText $errText -LogLevel "ERROR"
                    $results.Add([PSCustomObject]@{
                        SiteUrl      = $cleanUrl
                        Status       = "error"
                        ErrorMessage = $_.ToString()
                    })
                    $siteDone = $true
                }
            }
        }
    }

    # Summary
    $completed = ($results | Where-Object { $_.Status -eq "completed" }).Count
    $failed    = ($results | Where-Object { $_.Status -eq "failed" }).Count
    $errors    = ($results | Where-Object { $_.Status -eq "error" }).Count
    $other     = $results.Count - $completed - $failed - $errors

    Write-Host "`n==== Report Status Summary ====" -ForegroundColor Cyan
    Write-Host "  Total sites processed : $($results.Count)" -ForegroundColor White
    Write-Host "  Completed             : $completed" -ForegroundColor Green
    if ($failed -gt 0) {
        Write-Host "  Failed                : $failed" -ForegroundColor Red
        Write-Host "`n  Failed sites:" -ForegroundColor Red
        $results | Where-Object { $_.Status -eq "failed" } | ForEach-Object {
            Write-Host "    - $($_.SiteUrl)" -ForegroundColor Red
            Write-Host "      $($_.ErrorMessage)" -ForegroundColor DarkRed
        }
    }
    if ($errors -gt 0) {
        Write-Host "  Errors (cmdlet)       : $errors" -ForegroundColor Red
    }
    if ($other -gt 0) {
        Write-Host "  Other/In-Progress     : $other" -ForegroundColor Yellow
        $results | Where-Object { $_.Status -notin @("completed","failed","error") } | ForEach-Object {
            Write-Host "    - $($_.SiteUrl) : $($_.Status)" -ForegroundColor Yellow
        }
    }

    Write-LogEntry -LogName $log -LogEntryText "Report status summary: Total=$($results.Count), Completed=$completed, Failed=$failed, Errors=$errors, Other=$other" -LogLevel "INFO"
}

# Helper: applies a What-If policy to a downloaded CSV and returns storage impact metrics
function Get-WhatIfStorageAnalysis {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Automatic", "ExpireAfter", "CountLimit")]
        [string]$Mode,
        [double]$ExpireAfterDays = 0,
        [int]$MajorVersionLimit  = 0
    )

    $report = Import-Csv -Path $CsvPath
    if ($report.Count -eq 0) {
        return @{ TotalVersions = 0; VersionsToDelete = 0; StorageFreedBytes = 0
                  StorageFreedMB = 0; StorageFreedGB = 0; TotalVersionStorageMB = 0; PercentFreed = 0 }
    }

    # Expand compact columns — the CSV omits repeated values for the same file to save space
    $prevWebId = ""; $prevDocId = ""; $prevWebUrl = ""; $prevFileUrl = ""
    $prevModUser = ""; $prevModName = ""
    foreach ($row in $report) {
        if (![string]::IsNullOrEmpty($row."WebId.Compact"))                 { $prevWebId   = $row."WebId.Compact" }                  else { $row."WebId.Compact"                  = $prevWebId }
        if (![string]::IsNullOrEmpty($row."DocId.Compact"))                 { $prevDocId   = $row."DocId.Compact" }                  else { $row."DocId.Compact"                  = $prevDocId }
        if (![string]::IsNullOrEmpty($row."WebUrl.Compact"))                { $prevWebUrl  = $row."WebUrl.Compact" }                 else { $row."WebUrl.Compact"                 = $prevWebUrl }
        if (![string]::IsNullOrEmpty($row."FileUrl.Compact"))               { $prevFileUrl = $row."FileUrl.Compact" }                else { $row."FileUrl.Compact"                = $prevFileUrl }
        if (![string]::IsNullOrEmpty($row."ModifiedBy_UserId.Compact"))     { $prevModUser = $row."ModifiedBy_UserId.Compact" }      else { $row."ModifiedBy_UserId.Compact"      = $prevModUser }
        if (![string]::IsNullOrEmpty($row."ModifiedBy_DisplayName.Compact")){ $prevModName = $row."ModifiedBy_DisplayName.Compact" } else { $row."ModifiedBy_DisplayName.Compact" = $prevModName }
    }

    switch ($Mode) {
        "Automatic" {
            foreach ($row in $report) {
                $row.TargetExpirationDate = $row.AutomaticPolicyExpirationDate
            }
        }
        "ExpireAfter" {
            foreach ($row in $report) {
                if (![string]::IsNullOrEmpty($row.SnapshotDate)) {
                    try {
                        $snap = [DateTime]::Parse($row.SnapshotDate)
                        $row.TargetExpirationDate = $snap.AddDays($ExpireAfterDays).ToString("yyyy-MM-ddTHH:mm:ssK")
                    } catch { }
                }
            }
        }
        "CountLimit" {
            # Group versions per file, sort descending by version number, mark excess as expired
            $fileGroups = $report | Group-Object -Property "DocId.Compact"
            foreach ($group in $fileGroups) {
                $sorted = $group.Group | Sort-Object { [int]$_.MajorVersion * 512 + [int]$_.MinorVersion } -Descending
                $majorCount = 0
                foreach ($v in $sorted) {
                    if ($majorCount -ge $MajorVersionLimit) {
                        $v.TargetExpirationDate = "2000-01-01T00:00:00Z"
                    }
                    if ([int]$v.MinorVersion -eq 0) { $majorCount++ }
                }
            }
        }
    }

    $toDelete = $report | Where-Object { ![string]::IsNullOrEmpty($_.TargetExpirationDate) }

    $storageFreedBytes = [long]0
    $totalStorageBytes = [long]0
    foreach ($v in $report) {
        $sz = [long]0
        if (![string]::IsNullOrEmpty($v.Size) -and [long]::TryParse($v.Size, [ref]$sz)) {
            $totalStorageBytes += $sz
        }
    }
    foreach ($v in $toDelete) {
        $sz = [long]0
        if (![string]::IsNullOrEmpty($v.Size) -and [long]::TryParse($v.Size, [ref]$sz)) {
            $storageFreedBytes += $sz
        }
    }

    return @{
        TotalVersions         = $report.Count
        VersionsToDelete      = $toDelete.Count
        StorageFreedBytes     = $storageFreedBytes
        StorageFreedMB        = [math]::Round($storageFreedBytes / 1MB, 2)
        StorageFreedGB        = [math]::Round($storageFreedBytes / 1GB, 3)
        TotalVersionStorageMB = [math]::Round($totalStorageBytes / 1MB, 2)
        PercentFreed          = if ($totalStorageBytes -gt 0) { [math]::Round(($storageFreedBytes / $totalStorageBytes) * 100, 1) } else { 0 }
    }
}

# Function to run What-If analysis across all site collections
function Invoke-TenantVersionWhatIfAnalysis {
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$SiteUrls,
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $false)]
        [string]$TenantId
    )

    Write-Host "`n==== Version Policy What-If Analysis ====" -ForegroundColor Cyan
    Write-Host "Downloads version reports from each site and calculates how much storage" -ForegroundColor Yellow
    Write-Host "would be recovered under the selected version policy." -ForegroundColor Yellow
    Write-LogEntry -LogName $log -LogEntryText "Starting What-If analysis" -LogLevel "INFO"

    # Resolve site list
    if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) {
        Write-Host "`nNo sites loaded. Starting site discovery..." -ForegroundColor Yellow
        $scopeChoice = Get-SiteScope
        if ($scopeChoice -eq "3") { Write-Host "Operation cancelled by user." -ForegroundColor Yellow; return }
        $SiteUrls = Get-FilteredSites -Scope $scopeChoice
        if ($null -eq $SiteUrls -or $SiteUrls.Count -eq 0) { Write-Host "No sites found. Operation cancelled." -ForegroundColor Red; return }
        $confirm = Read-Host "`nReady to process $($SiteUrls.Count) sites. Proceed? (Y/N)"
        if ($confirm -ne "Y" -and $confirm -ne "y") { Write-Host "Operation cancelled by user." -ForegroundColor Yellow; return }
    }

    $filteredSites = Get-VersionReportEligibleSites -SiteUrls $SiteUrls -MinSiteSizeMB $script:MinSiteSizeforversionReports
    $SiteUrls = $filteredSites.EligibleSiteUrls
    if ($SiteUrls.Count -eq 0) {
        Write-Host "No sites meet the minimum size threshold ($($script:MinSiteSizeforversionReports) MB)." -ForegroundColor Yellow
        Write-LogEntry -LogName $log -LogEntryText "No sites eligible for What-If analysis after size filter. ThresholdMB=$($script:MinSiteSizeforversionReports)" -LogLevel "WARNING"
        return
    }

    # Choose policy mode
    Write-Host "`n==== Select Version Policy to Analyze ====" -ForegroundColor Cyan
    Write-Host "1: Automatic version trimming (uses AutomaticPolicyExpirationDate from report)"
    Write-Host "2: Manual - Expire versions older than X days"
    Write-Host "3: Manual - Keep only N most recent major versions"

    $modeChoice = $null
    do {
        $modeChoice = Read-Host "Select policy (1-3)"
        if ($modeChoice -notin @("1", "2", "3")) { Write-Host "Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red }
    } while ($modeChoice -notin @("1", "2", "3"))

    $analysisMode      = ""
    $expireAfterDays   = [double]0
    $majorVersionLimit = [int]0
    $policyDescription = ""

    switch ($modeChoice) {
        "1" {
            $analysisMode      = "Automatic"
            $policyDescription = "Automatic version trimming"
        }
        "2" {
            $analysisMode = "ExpireAfter"
            do {
                $daysInput = Read-Host "Enter number of days — versions older than this will be deleted (minimum 30)"
                $validDays = [double]::TryParse($daysInput, [ref]$expireAfterDays)
                if (-not $validDays -or $expireAfterDays -lt 30) { Write-Host "Please enter a number of 30 or greater." -ForegroundColor Red }
            } while (-not $validDays -or $expireAfterDays -lt 30)
            $policyDescription = "Manual expiration: versions older than $expireAfterDays days"
        }
        "3" {
            $analysisMode = "CountLimit"
            do {
                $countInput = Read-Host "Enter the number of most recent major versions to KEEP (minimum 1)"
                $validCount = [int]::TryParse($countInput, [ref]$majorVersionLimit)
                if (-not $validCount -or $majorVersionLimit -lt 1) { Write-Host "Please enter a positive integer." -ForegroundColor Red }
            } while (-not $validCount -or $majorVersionLimit -lt 1)
            $policyDescription = "Manual count limit: keep $majorVersionLimit most recent major versions"
        }
    }

    Write-Host "`nPolicy selected: $policyDescription" -ForegroundColor Green
    Write-LogEntry -LogName $log -LogEntryText "What-If mode: $analysisMode | Policy: $policyDescription" -LogLevel "INFO"

    # Create timestamped temp directory
    $tempDir = Join-Path $env:TEMP "SPO_WhatIf_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    Write-Host "Downloading reports to: $tempDir" -ForegroundColor Cyan
    Write-LogEntry -LogName $log -LogEntryText "Temp directory: $tempDir" -LogLevel "INFO"

    $reportLibraryName = "Admin_SiteCollection_VersionReport_DONOTDELETE"
    $siteResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($siteUrl in $SiteUrls) {
        $cleanUrl = $siteUrl.TrimEnd('/')
        Write-Host "`nProcessing: $cleanUrl" -ForegroundColor Cyan

        $siteCollectionName = ($cleanUrl -split '/') | Where-Object { $_ -ne '' } | Select-Object -Last 1
        $reportFileName = "${siteCollectionName}site_adminreport_donotdelete_VersionReport.csv"
        $localCsvPath   = Join-Path $tempDir $reportFileName

        $siteRetry = 0
        $siteDone = $false
        while (-not $siteDone) {
            try {
                Connect-PnPOnline -Url $cleanUrl -ClientId $ClientId -Tenant $TenantId -Interactive

                # Verify report file exists before attempting download
                $reportFile = Get-PnPFile -Url "/$reportLibraryName/$reportFileName" -ErrorAction SilentlyContinue
                if ($null -eq $reportFile) {
                    Write-Host "  - Report not found. Run option 9 first to generate reports." -ForegroundColor Yellow
                    Write-LogEntry -LogName $log -LogEntryText "Report not found for $cleanUrl — skipping" -LogLevel "WARNING"
                    break
                }

                Write-Host "  - Downloading report..." -ForegroundColor Cyan
                Get-PnPFile -Url "/$reportLibraryName/$reportFileName" -Path $tempDir -Filename $reportFileName -AsFile -Force

                Write-Host "  - Applying What-If analysis ($policyDescription)..." -ForegroundColor Cyan
                $analysis = Get-WhatIfStorageAnalysis -CsvPath $localCsvPath -Mode $analysisMode `
                    -ExpireAfterDays $expireAfterDays -MajorVersionLimit $majorVersionLimit

                Write-Host "  - Versions in report  : $($analysis.TotalVersions)" -ForegroundColor White
                Write-Host "  - Versions to delete  : $($analysis.VersionsToDelete)" -ForegroundColor Yellow
                Write-Host "  - Version storage used: $($analysis.TotalVersionStorageMB) MB" -ForegroundColor White
                Write-Host "  - Storage to recover  : $($analysis.StorageFreedMB) MB ($($analysis.StorageFreedGB) GB)" -ForegroundColor Green
                Write-Host "  - % storage recovered : $($analysis.PercentFreed)%" -ForegroundColor Green

                $siteResults.Add([PSCustomObject]@{
                    SiteUrl               = $cleanUrl
                    TotalVersions         = $analysis.TotalVersions
                    VersionsToDelete      = $analysis.VersionsToDelete
                    TotalVersionStorageMB = $analysis.TotalVersionStorageMB
                    StorageFreedMB        = $analysis.StorageFreedMB
                    StorageFreedGB        = $analysis.StorageFreedGB
                    PercentFreed          = $analysis.PercentFreed
                })

                Write-LogEntry -LogName $log -LogEntryText "What-If for $cleanUrl : Versions=$($analysis.TotalVersions), ToDelete=$($analysis.VersionsToDelete), TotalStorageMB=$($analysis.TotalVersionStorageMB), FreedMB=$($analysis.StorageFreedMB), Percent=$($analysis.PercentFreed)%" -LogLevel "INFO"
                $siteDone = $true
            }
            catch {
                $statusCode = $null
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
                if (($statusCode -eq 429 -or $statusCode -eq 503) -and $siteRetry -lt 5) {
                    $retryAfter = $null
                    try { $retryAfter = $_.Exception.Response.Headers["Retry-After"] } catch { }
                    if (-not $retryAfter) { $retryAfter = 30 * [math]::Pow(2, $siteRetry) }
                    $siteRetry++
                    $msg = "Throttling detected for $cleanUrl. Waiting $([int][math]::Ceiling($retryAfter))s before retry $siteRetry/5..."
                    Write-Warning $msg
                    Write-LogEntry -LogName $log -LogEntryText $msg -LogLevel "WARNING"
                    Start-Sleep -Seconds ([int][math]::Ceiling($retryAfter))
                }
                else {
                    $errMsg = "Failed What-If analysis for $cleanUrl : $_"
                    Write-Error $errMsg
                    Write-Host $_.Exception.ToString() -ForegroundColor Red
                    Write-LogEntry -LogName $log -LogEntryText $errMsg -LogLevel "ERROR"
                    $siteDone = $true
                }
            }
        }
    }

    # Aggregate summary
    if ($siteResults.Count -gt 0) {
        $totalVersions         = ($siteResults | Measure-Object -Property TotalVersions         -Sum).Sum
        $totalToDelete         = ($siteResults | Measure-Object -Property VersionsToDelete       -Sum).Sum
        $totalVersionStorageMB = [math]::Round(($siteResults | Measure-Object -Property TotalVersionStorageMB -Sum).Sum, 2)
        $totalFreedMB          = [math]::Round(($siteResults | Measure-Object -Property StorageFreedMB        -Sum).Sum, 2)
        $totalFreedGB          = [math]::Round($totalFreedMB / 1024, 3)
        $overallPct            = if ($totalVersionStorageMB -gt 0) { [math]::Round(($totalFreedMB / $totalVersionStorageMB) * 100, 1) } else { 0 }

        Write-Host "`n==== What-If Analysis Summary ====" -ForegroundColor Cyan
        Write-Host "  Policy analyzed         : $policyDescription" -ForegroundColor White
        Write-Host "  Sites analyzed          : $($siteResults.Count)" -ForegroundColor White
        Write-Host "  Total versions          : $totalVersions" -ForegroundColor White
        Write-Host "  Versions to delete      : $totalToDelete" -ForegroundColor Yellow
        Write-Host "  Total version storage   : $totalVersionStorageMB MB" -ForegroundColor White
        Write-Host "  Total storage to recover: $totalFreedMB MB  ($totalFreedGB GB)" -ForegroundColor Green
        Write-Host "  Overall % recovered     : $overallPct%" -ForegroundColor Green

        Write-Host "`n  Per-site breakdown (sorted by storage freed):" -ForegroundColor Cyan
        $siteResults | Sort-Object StorageFreedMB -Descending | ForEach-Object {
            Write-Host ("    {0,-55} {1,8} MB freed  ({2}%)" -f $_.SiteUrl, $_.StorageFreedMB, $_.PercentFreed) -ForegroundColor White
        }

        # Export results to CSV
        $safePolicyName  = $policyDescription -replace '[\\/:*?"<>|]', '_'
        $csvExportPath   = Join-Path $env:TEMP "SPO_WhatIf_Results_${safePolicyName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        $exportRows = $siteResults | Sort-Object StorageFreedMB -Descending |
            Select-Object SiteUrl, TotalVersions, VersionsToDelete,
                          TotalVersionStorageMB, StorageFreedMB, StorageFreedGB, PercentFreed

        # Append a totals row
        $totalsRow = [PSCustomObject]@{
            SiteUrl               = "TOTAL ($($siteResults.Count) sites)"
            TotalVersions         = $totalVersions
            VersionsToDelete      = $totalToDelete
            TotalVersionStorageMB = $totalVersionStorageMB
            StorageFreedMB        = $totalFreedMB
            StorageFreedGB        = $totalFreedGB
            PercentFreed          = $overallPct
        }
        $exportRows + $totalsRow | Export-Csv -Path $csvExportPath -NoTypeInformation

        Write-Host "`n  Results exported to: $csvExportPath" -ForegroundColor Green
        Write-LogEntry -LogName $log -LogEntryText "What-If summary: Policy=$policyDescription, Sites=$($siteResults.Count), TotalVersions=$totalVersions, ToDelete=$totalToDelete, TotalStorageMB=$totalVersionStorageMB, FreedMB=$totalFreedMB, FreedGB=$totalFreedGB, Percent=$overallPct%" -LogLevel "INFO"
        Write-LogEntry -LogName $log -LogEntryText "What-If results exported to: $csvExportPath" -LogLevel "INFO"
    }

    # Offer to keep or clean up temp files
    $keepFiles = Read-Host "`nKeep downloaded report files at '$tempDir'? (Y/N)"
    if ($keepFiles -ne "Y" -and $keepFiles -ne "y") {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Temporary files removed." -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "What-If temp files removed: $tempDir" -LogLevel "INFO"
    }
    else {
        Write-Host "Files retained at: $tempDir" -ForegroundColor Cyan
        Write-LogEntry -LogName $log -LogEntryText "What-If temp files retained at: $tempDir" -LogLevel "INFO"
    }
}

# Display menu and get user selection
function Show-OperationMenu {
    Clear-Host
    Write-Host "==== SharePoint Site Version Policy Operations ====" -ForegroundColor Cyan
    Write-Host ""
    
    # Display site discovery mode
    if ($null -ne $script:sitesFilePath -and $script:sitesFilePath -ne "" -and $null -ne $script:sites) {
        Write-Host "Site Mode: Batch processing ($($script:sites.Count) sites from file)" -ForegroundColor Green
    }
    elseif ($null -eq $script:sitesFilePath -or $script:sitesFilePath -eq "") {
        Write-Host "Site Mode: Auto-discovery (all sites in tenant)" -ForegroundColor Green
        Write-Host "  You will be prompted to select SharePoint or OneDrive sites" -ForegroundColor Gray
    }
    else {
        Write-Host "Site Mode: Batch processing (file configured but not loaded)" -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "Site-Level Operations:" -ForegroundColor Yellow
    Write-Host "1: Get current version policy for all sites"
    Write-Host "2: Set version policy for all sites"
    Write-Host "3: Get version policy status for all sites"
    Write-Host "4: Create batch delete job for all sites"
    Write-Host "5: Get batch delete job status for all sites"
    Write-Host ""
    Write-Host "Tenant-Level Operations (applies to new sites):" -ForegroundColor Yellow
    Write-Host "6: Set tenant to automatic version trimming"
    Write-Host "7: Set tenant to manual version limits"
    Write-Host "8: Review current tenant level version settings"
    Write-Host "9: Generate version history report for all sites"
    Write-Host "10: Get version history report job status for all sites"
    Write-Host "11: What-If analysis - estimate storage recovery by version policy"
    Write-Host ""
    Write-Host "Q: Quit"
    Write-Host "====================================================" -ForegroundColor Cyan
    
    $selection = Read-Host "Please select an operation (1-11, or Q to quit)"
    Write-LogEntry -LogName $log -LogEntryText "User selected menu option: $selection" -LogLevel "INFO"
    return $selection
}

# Main execution loop
function Start-OperationsMenu {
    $continue = $true
    Write-LogEntry -LogName $log -LogEntryText "Starting operations menu" -LogLevel "INFO"
    
    while ($continue) {
        $choice = Show-OperationMenu
        
        switch ($choice) {
            "1" {
                Write-Host "Running: Get current version policy" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get current version policy" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getVersionPolicyOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "get version policy"
                Read-Host "Press Enter to return to menu"
            }
            "2" {
                Write-Host "Running: Set version policy for all sites" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Set version policy for all sites" -LogLevel "INFO"
                
                # Ask user to choose between automatic or manual
                Write-Host "`n==== Choose Version Policy Type ====" -ForegroundColor Cyan
                Write-Host "1: Automatic (intelligent algorithm)"
                Write-Host "2: Manual (with version limits)"
                
                $policyTypeChoice = $null
                do {
                    $policyTypeChoice = Read-Host "Select policy type (1-2)"
                    if ($policyTypeChoice -notin @("1", "2")) {
                        Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                    }
                } while ($policyTypeChoice -notin @("1", "2"))
                
                if ($policyTypeChoice -eq "1") {
                    # Automatic mode
                    Write-Host "`nSetting automatic version trimming for all sites..." -ForegroundColor Cyan
                    Write-LogEntry -LogName $log -LogEntryText "User selected automatic version trimming" -LogLevel "INFO"
                    Invoke-SiteBatch -SiteUrls $sites -Operation $setVersionPolicyOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "set automatic version policy"
                }
                else {
                    # Manual mode - ask for settings source
                    Write-Host "`n==== Choose Settings Source ====" -ForegroundColor Cyan
                    Write-Host "1: Use tenant-level settings (apply current tenant defaults to all sites)"
                    Write-Host "2: Enter custom settings"
                    
                    $settingChoice = $null
                    do {
                        $settingChoice = Read-Host "Select option (1-2)"
                        if ($settingChoice -notin @("1", "2")) {
                            Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                        }
                    } while ($settingChoice -notin @("1", "2"))
                    
                    $manualSettings = $null
                    
                    if ($settingChoice -eq "1") {
                        # Get tenant-level settings
                        Write-Host "`nRetrieving tenant-level version settings..." -ForegroundColor Cyan
                        Write-LogEntry -LogName $log -LogEntryText "Retrieving tenant-level settings to apply to sites" -LogLevel "INFO"
                        
                        try {
                            $tenantConfig = Get-PnPTenant
                            
                            # Extract version settings from tenant
                            $manualSettings = @{
                                MajorVersionLimit = $tenantConfig.MajorVersionLimit
                                ExpireAfterDays   = if ($tenantConfig.ExpireVersionsAfterDays -eq 0) { $null } else { $tenantConfig.ExpireVersionsAfterDays }
                            }
                            
                            Write-Host "`nTenant-level settings retrieved:" -ForegroundColor Green
                            Write-Host "  Major Version Limit: $($manualSettings.MajorVersionLimit)" -ForegroundColor Green
                            if ($null -ne $manualSettings.ExpireAfterDays) {
                                Write-Host "  Expire After Days: $($manualSettings.ExpireAfterDays)" -ForegroundColor Green
                            }
                            else {
                                Write-Host "  Expire After Days: Never (No Expiration)" -ForegroundColor Green
                            }
                            
                            Write-LogEntry -LogName $log -LogEntryText "Retrieved tenant settings: MajorVersionLimit = $($manualSettings.MajorVersionLimit), ExpireAfterDays = $(if ($null -ne $manualSettings.ExpireAfterDays) { $manualSettings.ExpireAfterDays } else { 'Never' })" -LogLevel "INFO"
                            
                            $confirm = Read-Host "`nApply these settings to all sites? (Y/N)"
                            if ($confirm -ne "Y" -and $confirm -ne "y") {
                                Write-Host "Operation cancelled by user." -ForegroundColor Yellow
                                Write-LogEntry -LogName $log -LogEntryText "User cancelled operation after viewing tenant settings" -LogLevel "INFO"
                                Read-Host "Press Enter to return to menu"
                                continue
                            }
                        }
                        catch {
                            $errorMsg = "Failed to retrieve tenant settings: $_"
                            Write-Error $errorMsg
                            Write-Host $_.Exception.ToString() -ForegroundColor Red
                            Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                            Read-Host "Press Enter to return to menu"
                            continue
                        }
                    }
                    else {
                        # Get custom user input for manual version settings
                        $manualSettings = Get-ManualVersionSettings
                    }
                    
                    # Store settings in script scope for scriptblock access
                    $script:currentMajorVersionLimit = $manualSettings.MajorVersionLimit
                    $script:currentExpireAfterDays = $manualSettings.ExpireAfterDays
                    
                    # Create the operation scriptblock with the selected settings
                    $script:setManualVersionPolicyOperation = {
                        # Build parameters for Set-PnPSiteVersionPolicy
                        # When EnableAutoExpirationVersionTrim is false, ALL three parameters are required:
                        # MajorVersions, MajorWithMinorVersions, and ExpireVersionsAfterDays
                        $params = @{
                            EnableAutoExpirationVersionTrim = $false
                            MajorVersions                   = $script:currentMajorVersionLimit
                            MajorWithMinorVersions          = 0  # 0 means no minor versions kept
                        }
                        
                        # ExpireVersionsAfterDays is required - use 0 for "Never" if null
                        if ($null -ne $script:currentExpireAfterDays) {
                            $params.ExpireVersionsAfterDays = $script:currentExpireAfterDays
                        }
                        else {
                            $params.ExpireVersionsAfterDays = 0  # 0 means "Never expire"
                        }
                        
                        $result = Set-PnPSiteVersionPolicy @params
                        Write-Host "  - Site manual version policy set successfully" -ForegroundColor Green
                        Write-Host "    Major Version Limit: $($script:currentMajorVersionLimit)" -ForegroundColor Green
                        Write-Host "    Major with Minor Versions: 0 (no minor versions)" -ForegroundColor Green
                        if ($null -ne $script:currentExpireAfterDays) {
                            Write-Host "    Expire After Days: $($script:currentExpireAfterDays)" -ForegroundColor Green
                        }
                        else {
                            Write-Host "    Expire After Days: Never (No Expiration)" -ForegroundColor Green
                        }
                        Write-LogEntry -LogName $log -LogEntryText "Site manual version policy set: EnableAutoExpirationVersionTrim = False, MajorVersions = $($script:currentMajorVersionLimit), MajorWithMinorVersions = 0, ExpireAfterDays = $(if ($null -ne $script:currentExpireAfterDays) { $script:currentExpireAfterDays } else { '0 (Never)' })" -LogLevel "INFO"
                        return $result | Format-List
                    }
                    
                    # Execute the batch operation
                    Invoke-SiteBatch -SiteUrls $sites -Operation $setManualVersionPolicyOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "set manual version policy"
                }
                
                Read-Host "Press Enter to return to menu"
            }
            "3" {
                Write-Host "Running: Get version policy status" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get version policy status" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getVersionPolicyStatusOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "get version policy status"
                Read-Host "Press Enter to return to menu"
            }
            "4" {
                Write-Host "Running: Create batch delete job" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Create batch delete job" -LogLevel "INFO"
                
                # Ask user to choose between automatic or manual
                Write-Host "`n==== Choose Batch Delete Mode ====" -ForegroundColor Cyan
                Write-Host "1: Automatic (based on current site version policy)"
                Write-Host "2: Manual (with custom deletion settings)"
                
                $deleteMode = $null
                do {
                    $deleteMode = Read-Host "Select deletion mode (1-2)"
                    if ($deleteMode -notin @("1", "2")) {
                        Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                    }
                } while ($deleteMode -notin @("1", "2"))
                
                if ($deleteMode -eq "1") {
                    # Automatic mode - uses current site version policy
                    # First, retrieve and display tenant-level settings for confirmation
                    Write-Host "`nRetrieving tenant-level version settings..." -ForegroundColor Cyan
                    Write-LogEntry -LogName $log -LogEntryText "User selected automatic batch delete mode - retrieving tenant settings" -LogLevel "INFO"
                    
                    try {
                        $tenantConfig = Get-PnPTenant
                        
                        Write-Host "`n==== Current Tenant Version Settings ====" -ForegroundColor Cyan
                        Write-Host "Automatic batch delete will use the current site version policies," -ForegroundColor Yellow
                        Write-Host "which are based on these tenant-level defaults:" -ForegroundColor Yellow
                        Write-Host ""
                        
                        # Display tenant settings
                        if ($tenantConfig.EnableAutoExpirationVersionTrim -eq $true) {
                            Write-Host "  Mode: Automatic Version Trimming" -ForegroundColor Green
                            Write-Host "  Description: Uses intelligent algorithm to optimize storage" -ForegroundColor Green
                        }
                        else {
                            Write-Host "  Mode: Manual Version Limits" -ForegroundColor Green
                            Write-Host "  Major Version Limit: $($tenantConfig.MajorVersionLimit)" -ForegroundColor Green
                            if ($tenantConfig.ExpireVersionsAfterDays -eq 0) {
                                Write-Host "  Expire After Days: Never (No Expiration)" -ForegroundColor Green
                            }
                            else {
                                Write-Host "  Expire After Days: $($tenantConfig.ExpireVersionsAfterDays)" -ForegroundColor Green
                            }
                        }
                        
                        Write-Host ""
                        Write-Host "Note: Each site will use its own version policy for automatic deletion." -ForegroundColor Yellow
                        Write-Host "Sites not yet configured will use the tenant defaults shown above." -ForegroundColor Yellow
                        
                        Write-LogEntry -LogName $log -LogEntryText "Displayed tenant settings: EnableAutoExpirationVersionTrim = $($tenantConfig.EnableAutoExpirationVersionTrim), MajorVersionLimit = $($tenantConfig.MajorVersionLimit), ExpireVersionsAfterDays = $($tenantConfig.ExpireVersionsAfterDays)" -LogLevel "INFO"
                        
                        $confirm = Read-Host "`nProceed with automatic batch delete for all sites? (Y/N)"
                        if ($confirm -ne "Y" -and $confirm -ne "y") {
                            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
                            Write-LogEntry -LogName $log -LogEntryText "User cancelled automatic batch delete operation" -LogLevel "INFO"
                            Read-Host "Press Enter to return to menu"
                            continue
                        }
                        
                        Write-Host "`nCreating automatic batch delete jobs for all sites..." -ForegroundColor Cyan
                        Invoke-SiteBatch -SiteUrls $sites -Operation $createBatchDeleteJobOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "create automatic batch delete job"
                    }
                    catch {
                        $errorMsg = "Failed to retrieve tenant settings: $_"
                        Write-Error $errorMsg
                        Write-Host $_.Exception.ToString() -ForegroundColor Red
                        Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                        Read-Host "Press Enter to return to menu"
                        continue
                    }
                }
                else {
                    # Manual mode - ask for settings source
                    Write-Host "`n==== Choose Settings Source ====" -ForegroundColor Cyan
                    Write-Host "1: Use tenant-level settings (apply current tenant defaults to all sites)"
                    Write-Host "2: Enter custom settings"
                    
                    $settingChoice = $null
                    do {
                        $settingChoice = Read-Host "Select option (1-2)"
                        if ($settingChoice -notin @("1", "2")) {
                            Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                        }
                    } while ($settingChoice -notin @("1", "2"))
                    
                    $deleteSettings = $null
                    
                    if ($settingChoice -eq "1") {
                        # Get tenant-level settings
                        Write-Host "`nRetrieving tenant-level version settings..." -ForegroundColor Cyan
                        Write-LogEntry -LogName $log -LogEntryText "Retrieving tenant-level settings for batch delete" -LogLevel "INFO"
                        
                        try {
                            $tenantConfig = Get-PnPTenant
                            
                            # Build delete settings based on tenant configuration
                            # Check if tenant has MajorVersionLimit and ExpireVersionsAfterDays set
                            if ($tenantConfig.MajorVersionLimit -gt 0 -and $tenantConfig.ExpireVersionsAfterDays -gt 0) {
                                # Both settings are configured - user must choose one
                                Write-Host "`nTenant-level settings retrieved:" -ForegroundColor Green
                                Write-Host "  Major Version Limit: $($tenantConfig.MajorVersionLimit)" -ForegroundColor Green
                                Write-Host "  Expire After Days: $($tenantConfig.ExpireVersionsAfterDays)" -ForegroundColor Green
                                Write-Host ""
                                Write-Host "Note: Batch delete job can only use ONE of these settings at a time." -ForegroundColor Yellow
                                Write-Host ""
                                Write-Host "==== Choose Which Setting to Use ====" -ForegroundColor Cyan
                                Write-Host "1: Use Major Version Limit ($($tenantConfig.MajorVersionLimit) versions)"
                                Write-Host "2: Use Expire After Days ($($tenantConfig.ExpireVersionsAfterDays) days)"
                                
                                $limitChoice = $null
                                do {
                                    $limitChoice = Read-Host "Select setting to use (1-2)"
                                    if ($limitChoice -notin @("1", "2")) {
                                        Write-Host "Invalid selection. Please choose 1 or 2." -ForegroundColor Red
                                    }
                                } while ($limitChoice -notin @("1", "2"))
                                
                                if ($limitChoice -eq "1") {
                                    $deleteSettings = @{
                                        MajorVersionLimit = $tenantConfig.MajorVersionLimit
                                    }
                                    Write-Host "`nUsing Major Version Limit: $($deleteSettings.MajorVersionLimit)" -ForegroundColor Green
                                    Write-LogEntry -LogName $log -LogEntryText "User selected MajorVersionLimit = $($deleteSettings.MajorVersionLimit) for batch delete" -LogLevel "INFO"
                                }
                                else {
                                    $deleteSettings = @{
                                        DeleteBeforeDays = $tenantConfig.ExpireVersionsAfterDays
                                    }
                                    Write-Host "`nUsing Expire After Days: $($deleteSettings.DeleteBeforeDays)" -ForegroundColor Green
                                    Write-LogEntry -LogName $log -LogEntryText "User selected DeleteBeforeDays = $($deleteSettings.DeleteBeforeDays) for batch delete" -LogLevel "INFO"
                                }
                            }
                            elseif ($tenantConfig.MajorVersionLimit -gt 0) {
                                $deleteSettings = @{
                                    MajorVersionLimit = $tenantConfig.MajorVersionLimit
                                }
                                
                                Write-Host "`nTenant-level settings retrieved:" -ForegroundColor Green
                                Write-Host "  Major Version Limit: $($deleteSettings.MajorVersionLimit)" -ForegroundColor Green
                            }
                            elseif ($tenantConfig.ExpireVersionsAfterDays -gt 0) {
                                $deleteSettings = @{
                                    DeleteBeforeDays = $tenantConfig.ExpireVersionsAfterDays
                                }
                                
                                Write-Host "`nTenant-level settings retrieved:" -ForegroundColor Green
                                Write-Host "  Expire After Days: $($deleteSettings.DeleteBeforeDays)" -ForegroundColor Green
                            }
                            else {
                                Write-Host "`nWarning: Tenant has no version expiration settings configured." -ForegroundColor Yellow
                                Write-Host "Using default: Delete versions older than 30 days" -ForegroundColor Yellow
                                $deleteSettings = @{
                                    DeleteOlderThanDays = 30
                                }
                            }
                            
                            Write-LogEntry -LogName $log -LogEntryText "Retrieved tenant delete settings: $(if ($deleteSettings.MajorVersionLimit) { "MajorVersionLimit = $($deleteSettings.MajorVersionLimit)" } elseif ($deleteSettings.DeleteBeforeDays) { "DeleteBeforeDays = $($deleteSettings.DeleteBeforeDays)" } else { "DeleteOlderThanDays = $($deleteSettings.DeleteOlderThanDays)" })" -LogLevel "INFO"
                            
                            $confirm = Read-Host "`nApply these settings to all sites? (Y/N)"
                            if ($confirm -ne "Y" -and $confirm -ne "y") {
                                Write-Host "Operation cancelled by user." -ForegroundColor Yellow
                                Write-LogEntry -LogName $log -LogEntryText "User cancelled batch delete operation" -LogLevel "INFO"
                                Read-Host "Press Enter to return to menu"
                                continue
                            }
                        }
                        catch {
                            $errorMsg = "Failed to retrieve tenant settings: $_"
                            Write-Error $errorMsg
                            Write-Host $_.Exception.ToString() -ForegroundColor Red
                            Write-LogEntry -LogName $log -LogEntryText $errorMsg -LogLevel "ERROR"
                            Read-Host "Press Enter to return to menu"
                            continue
                        }
                    }
                    else {
                        # Get custom user input for batch delete settings
                        $deleteSettings = Get-BatchDeleteSettings
                    }
                    
                    # Store settings in script scope for scriptblock access
                    $script:currentDeleteOlderThanDays = if ($deleteSettings.DeleteOlderThanDays) { $deleteSettings.DeleteOlderThanDays } else { $null }
                    $script:currentDeleteMajorVersionLimit = if ($deleteSettings.MajorVersionLimit) { $deleteSettings.MajorVersionLimit } else { $null }
                    $script:currentDeleteBeforeDays = if ($deleteSettings.DeleteBeforeDays) { $deleteSettings.DeleteBeforeDays } else { $null }
                    
                    # Create the operation scriptblock with the selected settings
                    $script:createManualBatchDeleteJobOperation = {
                        # Build parameters for New-PnPSiteFileVersionBatchDeleteJob
                        # Note: Only ONE parameter type can be used at a time (different parameter sets)
                        $params = @{
                            Force = $true
                        }
                        
                        # Add parameters based on what settings we have (only one will be set)
                        if ($script:currentDeleteOlderThanDays) {
                            $params.DeleteOlderThanDays = $script:currentDeleteOlderThanDays
                            Write-Host "  - Creating manual batch delete job (DeleteOlderThanDays: $($script:currentDeleteOlderThanDays))" -ForegroundColor Cyan
                        }
                        elseif ($script:currentDeleteMajorVersionLimit) {
                            # When using MajorVersionLimit, MajorWithMinorVersionsLimit is also required (Example 4)
                            $params.MajorVersionLimit = $script:currentDeleteMajorVersionLimit
                            $params.MajorWithMinorVersionsLimit = 0  # 0 means no minor versions kept
                            Write-Host "  - Creating manual batch delete job (MajorVersionLimit: $($script:currentDeleteMajorVersionLimit), MajorWithMinorVersionsLimit: 0)" -ForegroundColor Cyan
                        }
                        elseif ($script:currentDeleteBeforeDays) {
                            $params.DeleteBeforeDays = $script:currentDeleteBeforeDays
                            Write-Host "  - Creating manual batch delete job (DeleteBeforeDays: $($script:currentDeleteBeforeDays))" -ForegroundColor Cyan
                        }
                        
                        $job = New-PnPSiteFileVersionBatchDeleteJob @params
                        Write-Host "  - Site batch delete job created successfully" -ForegroundColor Green
                        Write-LogEntry -LogName $log -LogEntryText "Manual batch delete job created with settings: $(if ($script:currentDeleteOlderThanDays) { "DeleteOlderThanDays = $($script:currentDeleteOlderThanDays)" } elseif ($script:currentDeleteMajorVersionLimit) { "MajorVersionLimit = $($script:currentDeleteMajorVersionLimit), MajorWithMinorVersionsLimit = 0" } else { "DeleteBeforeDays = $($script:currentDeleteBeforeDays)" })" -LogLevel "INFO"
                        return $job | Format-List
                    }
                    
                    # Execute the batch operation
                    Invoke-SiteBatch -SiteUrls $sites -Operation $createManualBatchDeleteJobOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "create manual batch delete job"
                }
                
                Read-Host "Press Enter to return to menu"
            }
            "5" {
                Write-Host "Running: Get batch delete job status" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get batch delete job status" -LogLevel "INFO"
                Invoke-SiteBatch -SiteUrls $sites -Operation $getBatchDeleteJobStatusOperation -ClientId $clientId -TenantId $tenantId -Connection $connection -OperationDescription "get batch delete job status"
                Read-Host "Press Enter to return to menu"
            }
            "6" {
                Write-Host "Running: Set tenant to automatic version trimming" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Set tenant to automatic version trimming" -LogLevel "INFO"
                
                $result = Set-TenantAutomaticVersionPolicy -AdminUrl $url -ClientId $clientId -TenantId $tenantId
                
                if ($result) {
                    Write-Host "`nOperation completed successfully!" -ForegroundColor Green
                }
                else {
                    Write-Host "`nOperation failed. Check the log for details." -ForegroundColor Red
                }
                Read-Host "Press Enter to return to menu"
            }
            "7" {
                Write-Host "Running: Set tenant to manual version limits" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Set tenant to manual version limits" -LogLevel "INFO"
                
                $result = Set-TenantManualVersionPolicy -AdminUrl $url -ClientId $clientId -TenantId $tenantId
                
                if ($result) {
                    Write-Host "`nOperation completed successfully!" -ForegroundColor Green
                }
                else {
                    Write-Host "`nOperation failed. Check the log for details." -ForegroundColor Red
                }
                Read-Host "Press Enter to return to menu"
            }
            "8" {
                Write-Host "Running: Review current tenant level version settings" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Review tenant version settings" -LogLevel "INFO"
                
                $result = Get-TenantVersionSettings -AdminUrl $url -ClientId $clientId -TenantId $tenantId
                
                if (-not $result) {
                    Write-Host "`nFailed to retrieve settings. Check the log for details." -ForegroundColor Red
                }
                Read-Host "Press Enter to return to menu"
            }
            "9" {
                Write-Host "Running: Generate version history report for all sites" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Generate version expiration report for all sites" -LogLevel "INFO"
                Write-Host "Version report size threshold (MinSiteSizeforversionReports): $($script:MinSiteSizeforversionReports) MB" -ForegroundColor Cyan
                Write-LogEntry -LogName $log -LogEntryText "Version report size threshold for option 9: $($script:MinSiteSizeforversionReports) MB" -LogLevel "INFO"
                
                New-TenantVersionExpirationReport -SiteUrls $sites -ClientId $clientId -TenantId $tenantId
                
                Read-Host "Press Enter to return to menu"
            }
            "10" {
                Write-Host "Running: Get version history report job status for all sites" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: Get version expiration report job status for all sites" -LogLevel "INFO"
                Write-Host "Version report size threshold (MinSiteSizeforversionReports): $($script:MinSiteSizeforversionReports) MB" -ForegroundColor Cyan
                Write-LogEntry -LogName $log -LogEntryText "Version report size threshold for option 10: $($script:MinSiteSizeforversionReports) MB" -LogLevel "INFO"
                
                Get-TenantVersionExpirationReportStatus -SiteUrls $sites -ClientId $clientId -TenantId $tenantId
                
                Read-Host "Press Enter to return to menu"
            }
            "11" {
                Write-Host "Running: What-If analysis - estimate storage recovery by version policy" -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "Starting operation: What-If analysis" -LogLevel "INFO"
                Write-Host "Version report size threshold (MinSiteSizeforversionReports): $($script:MinSiteSizeforversionReports) MB" -ForegroundColor Cyan
                Write-LogEntry -LogName $log -LogEntryText "Version report size threshold for option 11: $($script:MinSiteSizeforversionReports) MB" -LogLevel "INFO"
                
                Invoke-TenantVersionWhatIfAnalysis -SiteUrls $sites -ClientId $clientId -TenantId $tenantId
                
                Read-Host "Press Enter to return to menu"
            }
            "Q" {
                $continue = $false
                Write-Host "Exiting script..." -ForegroundColor Yellow
                Write-LogEntry -LogName $log -LogEntryText "User exited script" -LogLevel "INFO"
            }
      
            default {
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                Write-LogEntry -LogName $log -LogEntryText "Invalid menu selection: $choice" -LogLevel "WARNING"
                Start-Sleep -Seconds 2
            }
        }
    }
}

# Start the interactive menu
Write-LogEntry -LogName $log -LogEntryText "Displaying operations menu" -LogLevel "INFO"
Start-OperationsMenu

# Log script completion
Write-LogEntry -LogName $log -LogEntryText "Script execution completed. Log file: $log" -LogLevel "INFO"
write-host "Script execution completed. Log file: $log" -ForegroundColor Green
