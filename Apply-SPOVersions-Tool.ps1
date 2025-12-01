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
    Authors: Mike Lee /Luis DuSolier
    Date: 11/24/25

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
$Debug = $true

# This is the logging function
Function Write-LogEntry {
    param(
        [string] $LogName,
        [string] $LogEntryText,
        [string] $LogLevel = "INFO"  # Default log level is INFO
    )
    if ($LogName -ne $null) {
        # Skip DEBUG level messages if Debug is set to False
        if ($LogLevel -eq "DEBUG" -and $Debug -eq $False) {
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

$sitesFilePath = "C:\temp\M365CPI13246019-Sites.txt"  # Set to $null to auto-discover all sites
#$sitesFilePath = $null # Set to $null to auto-discover all sites

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
                $_.Template -notlike 'SRCHCEN*' -and
                $_.Template -notlike 'SRCHCENTERLITE*' -and
                $_.Template -notlike 'SPSMSITEHOST*' -and
                $_.Template -notlike 'APPCATALOG*' -and
                $_.Template -notlike 'REDIRECTSITE*' -and
                $_.Url -notlike '*-my.sharepoint.com/personal/*'
            }
            
            $siteUrls = $allSites | Select-Object -ExpandProperty Url
            Write-Host "Found $($siteUrls.Count) SharePoint sites" -ForegroundColor Green
            Write-LogEntry -LogName $log -LogEntryText "Found $($siteUrls.Count) SharePoint sites" -LogLevel "INFO"
        }
        elseif ($Scope -eq "2") {
            # Get only OneDrive for Business sites
            Write-Host "Retrieving OneDrive for Business sites..." -ForegroundColor Cyan
            Write-LogEntry -LogName $log -LogEntryText "Retrieving OneDrive sites" -LogLevel "INFO"
            
            $allSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"
            
            $siteUrls = $allSites | Select-Object -ExpandProperty Url
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

# Log script start
Write-LogEntry -LogName $log -LogEntryText "Script execution started. Connecting to tenant admin site: $url" -LogLevel "INFO"

# Connect to the SharePoint Online admin site
Connect-PnPOnline -Url $url -ClientId $clientId -Tenant $tenantId -Interactive
Write-LogEntry -LogName $log -LogEntryText "Successfully connected to admin site" -LogLevel "INFO"

# Load or discover sites based on configuration
$sites = $null

if ($null -ne $sitesFilePath -and $sitesFilePath -ne "") {
    # Load sites from file
    if (Test-Path $sitesFilePath) {
        $sites = Get-Content -Path $sitesFilePath
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
            if ($_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Response.StatusCode -eq 503) {
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                if (-not $retryAfter) {
                    $retryAfter = $InitialRetrySeconds * [math]::Pow(2, $retryCount)
                }
                
                $retryCount++
                $warningMsg = "Throttling detected for site $SiteUrl. Waiting for $retryAfter seconds before retry $retryCount of $MaxRetries..."
                Write-Warning $warningMsg
                Write-LogEntry -LogName $log -LogEntryText $warningMsg -LogLevel "WARNING"
                Start-Sleep -Seconds $retryAfter
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
        Set-PnPTenant -EnableAutoExpirationVersionTrim $true
        
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
        
        $tenantConfig = Get-PnPTenant
        
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
        Set-PnPTenant @params
        
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
    Write-Host ""
    Write-Host "Q: Quit"
    Write-Host "====================================================" -ForegroundColor Cyan
    
    $selection = Read-Host "Please select an operation (1-8, or Q to quit)"
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
