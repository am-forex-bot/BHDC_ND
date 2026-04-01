#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Wallace Office Startup Cleanup Script
    Fixes common causes of slow Word, Excel, and Outlook startup.

.DESCRIPTION
    Run this as admin on a user's machine (while they are logged in) to fix
    Office startup slowness. The script auto-detects the logged-in user and
    writes to their registry hive via HKU\<SID>.

    What it does:
    1. Removes dead/unwanted Outlook COM add-ins (iManage, ICQ, Lync, etc.)
    2. Clears the .NET assembly download cache (dl3) - stale DLLs cause slow startup
    3. Disables .NET Fusion assembly binding logging if enabled (massive perf killer)
    4. Cleans up iManage/Interwoven registry remnants
    5. Forces Outlook to keep required add-ins enabled (GPO-level resiliency override)
    6. Clears Outlook resiliency data (disabled/crashed add-in lists)
    7. Disables Word start screen and Office animations for faster perceived startup

.PARAMETER UserName
    Override the target username. Defaults to the currently logged-in user.

.PARAMETER WhatIf
    Shows what would be changed without making any changes.

.EXAMPLE
    .\Invoke-OfficeCleanup.ps1
    Runs cleanup for the currently logged-in user.

.EXAMPLE
    .\Invoke-OfficeCleanup.ps1 -UserName "john.smith"
    Runs cleanup for a specific user.

.EXAMPLE
    .\Invoke-OfficeCleanup.ps1 -WhatIf
    Shows what would be changed without making changes.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$UserName
)

$ErrorActionPreference = "Continue"

# --- Configuration -----------------------------------------------------------

# Add-ins to REMOVE from Outlook (dead/legacy/unwanted)
$AddinsToRemove = @(
    "ADXForm",
    "ColleagueImport.ColleagueImportAddin",
    "EmailScanner.Connect",
    "ICQExpress.CExpressClient.1",
    "Microsoft.OutlookBackup.1",
    "Microsoft.VbaAddinForOutlook.1",
    "OscAddin.Connect",
    "Redemption.Addin",
    "SBCM.Addin.1",
    "UCAddin.LyncAddin.1",
    "UmOutlookAddin.FormRegionAddin",
    "WorkSiteEmailManagement.Connect",
    "imFileSite.Connect",
    "OneNote.OutlookAddin"
)

# Add-ins to force-enable via GPO-level policy (Outlook cannot override these)
$AddinsToProtect = @(
    "Workshare.OutlookRibbon.Addin",
    "zzzDocsCorp.pdfDocs.OutlookAddIn",
    "MetaCompliance.Reporter.Email",
    "MimecastServicesForOutlook.Connect",
    "zzz.SafeSend",
    "TeamsAddin.FastConnect",
    "NetDocuments.Client.OutlookAddIn"
)

# Registry paths where Outlook add-ins can be registered
$OutlookAddinPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins"
)

# Fusion logging registry paths
$FusionLogPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Fusion",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Fusion"
)

# --- Helper Functions --------------------------------------------------------

function Write-Status {
    param([string]$Message, [string]$Type = "Info")
    $color = switch ($Type) {
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Skip"    { "DarkGray" }
        default   { "Cyan" }
    }
    Write-Host "  $Message" -ForegroundColor $color
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor White
}

# --- Resolve target user and their SID ---------------------------------------

if ($UserName) {
    $resolvedUser = $UserName
    $userProfile = "C:\Users\$UserName"
} else {
    $loggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
    if ($loggedInUser -and $loggedInUser -match "\\(.+)$") {
        $resolvedUser = $Matches[1]
        $userProfile = "C:\Users\$resolvedUser"
    } else {
        Write-Host "ERROR: Could not detect logged-in user. Use -UserName parameter." -ForegroundColor Red
        exit 1
    }
}

if (-not (Test-Path $userProfile)) {
    Write-Host "ERROR: User profile not found at $userProfile" -ForegroundColor Red
    exit 1
}

# Resolve SID for HKU access
try {
    $domainUser = if ($loggedInUser -and -not $UserName) { $loggedInUser } else { $resolvedUser }
    $userSID = (New-Object System.Security.Principal.NTAccount($domainUser)).Translate(
        [System.Security.Principal.SecurityIdentifier]).Value
    $hku = "Registry::HKU\$userSID"

    if (-not (Test-Path $hku)) {
        Write-Host "ERROR: HKU hive not accessible for SID $userSID. Is the user logged in?" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "ERROR: Could not resolve SID for $domainUser. Is the user logged in?" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Wallace Office Startup Cleanup" -ForegroundColor White
Write-Host "Target user: $resolvedUser ($userSID)" -ForegroundColor Cyan
Write-Host "Profile path: $userProfile" -ForegroundColor Cyan
Write-Host "Registry hive: $hku" -ForegroundColor Cyan
if ($WhatIfPreference) {
    Write-Host "*** WHATIF MODE - no changes will be made ***" -ForegroundColor Yellow
}

# --- Check for running Office apps -------------------------------------------

Write-Section "Checking for running Office applications"

$officeProcs = Get-Process OUTLOOK, WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue
if ($officeProcs) {
    Write-Host ""
    Write-Host "  WARNING: The following Office apps are running:" -ForegroundColor Yellow
    $officeProcs | ForEach-Object { Write-Host "    - $($_.Name) (PID: $($_.Id))" -ForegroundColor Yellow }
    Write-Host ""
    $response = Read-Host "  Close them now? (Y/N)"
    if ($response -eq "Y" -or $response -eq "y") {
        $officeProcs | Stop-Process -Force
        Write-Status "Office applications closed. Waiting for handles to release..." "Success"
        Start-Sleep -Seconds 5
    } else {
        Write-Host ""
        Write-Host "  Some cleanup tasks may fail with Office running." -ForegroundColor Yellow
        Write-Host "  Continuing anyway..." -ForegroundColor Yellow
    }
} else {
    Write-Status "No Office applications running." "Success"
}

# --- 1. Remove unwanted Outlook add-ins --------------------------------------

Write-Section "1. Removing unwanted Outlook add-ins"

$removedCount = 0
foreach ($addin in $AddinsToRemove) {
    # Check HKLM paths
    foreach ($basePath in $OutlookAddinPaths) {
        $fullPath = Join-Path $basePath $addin
        if (Test-Path $fullPath) {
            if ($PSCmdlet.ShouldProcess($fullPath, "Remove add-in")) {
                Remove-Item -Path $fullPath -Force -ErrorAction SilentlyContinue
                if (-not (Test-Path $fullPath)) {
                    Write-Status "Removed: $addin (from $basePath)" "Success"
                    $removedCount++
                } else {
                    Write-Status "Failed to remove: $addin (access denied?)" "Error"
                }
            }
        }
    }
    # Check user's HKCU via HKU
    $hkcuPath = "$hku\SOFTWARE\Microsoft\Office\Outlook\Addins\$addin"
    if (Test-Path $hkcuPath) {
        if ($PSCmdlet.ShouldProcess($hkcuPath, "Remove add-in")) {
            Remove-Item -Path $hkcuPath -Force -ErrorAction SilentlyContinue
            Write-Status "Removed: $addin (from user hive)" "Success"
            $removedCount++
        }
    }
}

if ($removedCount -eq 0) {
    Write-Status "No unwanted add-ins found - already clean." "Skip"
} else {
    Write-Status "Removed $removedCount add-in registrations." "Success"
}

# --- 2. Clear .NET assembly download cache -----------------------------------

Write-Section "2. Clearing .NET assembly download cache (dl3)"

$dl3Path = Join-Path $userProfile "AppData\Local\assembly\dl3"
if (Test-Path $dl3Path) {
    $dl3Size = (Get-ChildItem $dl3Path -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    $dl3SizeMB = [math]::Round($dl3Size / 1MB, 1)
    Write-Status "Found dl3 cache: $dl3SizeMB MB" "Info"

    if ($PSCmdlet.ShouldProcess($dl3Path, "Delete assembly cache")) {
        Remove-Item -Path $dl3Path -Recurse -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path $dl3Path)) {
            Write-Status "Cleared dl3 cache ($dl3SizeMB MB freed)." "Success"
        } else {
            Write-Status "Partial clear - some files were locked (Office still running?)." "Warning"
        }
    }
} else {
    Write-Status "No dl3 cache found - already clean." "Skip"
}

# --- 3. Disable Fusion assembly binding logging ------------------------------

Write-Section "3. Checking .NET Fusion logging"

$fusionFixed = $false
foreach ($fusionPath in $FusionLogPaths) {
    $fusionProps = Get-ItemProperty $fusionPath -ErrorAction SilentlyContinue
    if ($fusionProps) {
        $logEnabled = $false
        if ($fusionProps.ForceLog -eq 1) { $logEnabled = $true }
        if ($fusionProps.LogFailures -eq 1) { $logEnabled = $true }
        if ($fusionProps.EnableLog -eq 1) { $logEnabled = $true }
        if ($fusionProps.LogPath) { $logEnabled = $true }

        if ($logEnabled) {
            Write-Status "Fusion logging ENABLED at $fusionPath - this is a major perf killer!" "Warning"
            if ($PSCmdlet.ShouldProcess($fusionPath, "Disable Fusion logging")) {
                Set-ItemProperty $fusionPath -Name "ForceLog" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty $fusionPath -Name "LogFailures" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty $fusionPath -Name "EnableLog" -Value 0 -Type DWord -ErrorAction SilentlyContinue
                Remove-ItemProperty $fusionPath -Name "LogPath" -ErrorAction SilentlyContinue
                Write-Status "Disabled Fusion logging at $fusionPath" "Success"
                $fusionFixed = $true
            }
        }
    }
}

if (-not $fusionFixed) {
    Write-Status "Fusion logging not enabled - OK." "Skip"
}

# Clean up Fusion log output directory if it exists
$fusionLogDir = "C:\FusionLogs"
if (Test-Path $fusionLogDir) {
    $fusionLogSize = (Get-ChildItem $fusionLogDir -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    $fusionLogSizeMB = [math]::Round($fusionLogSize / 1MB, 1)
    Write-Status "Found Fusion log output: $fusionLogDir ($fusionLogSizeMB MB)" "Warning"
    if ($PSCmdlet.ShouldProcess($fusionLogDir, "Delete Fusion log files")) {
        Remove-Item -Path $fusionLogDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Status "Deleted Fusion log files." "Success"
    }
}

# --- 4. Clean up iManage/Interwoven remnants ---------------------------------

Write-Section "4. Cleaning up iManage/Interwoven registry remnants"

$imanagePaths = @(
    "HKLM:\SOFTWARE\Interwoven",
    "HKLM:\SOFTWARE\WOW6432Node\Interwoven",
    "$hku\SOFTWARE\Interwoven",
    "HKLM:\SOFTWARE\iManage",
    "HKLM:\SOFTWARE\WOW6432Node\iManage",
    "$hku\SOFTWARE\iManage"
)

$imanageCleaned = $false
foreach ($path in $imanagePaths) {
    if (Test-Path $path) {
        if ($PSCmdlet.ShouldProcess($path, "Remove iManage registry key")) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            if (-not (Test-Path $path)) {
                Write-Status "Removed: $path" "Success"
                $imanageCleaned = $true
            } else {
                Write-Status "Could not remove $path (access denied - may need SCCM ACL fix)" "Warning"
            }
        }
    }
}

if (-not $imanageCleaned) {
    Write-Status "No iManage remnants found." "Skip"
}

# --- 5. Protect required add-ins (GPO-level enforcement) --------------------

Write-Section "5. Protecting required Outlook add-ins from auto-disable"

# Use the Policies path - Outlook treats this as Group Policy and CANNOT override it
$policyPath = "$hku\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList"

if ($PSCmdlet.ShouldProcess($policyPath, "Set GPO-level add-in protection")) {
    New-Item -Path $policyPath -Force -ErrorAction SilentlyContinue | Out-Null
    foreach ($addin in $AddinsToProtect) {
        Set-ItemProperty $policyPath -Name $addin -Value 1 -Type DWord -ErrorAction SilentlyContinue
        Write-Status "Protected: $addin" "Success"
    }
}

# --- 6. Clear Outlook resiliency data (disabled/crashed add-in lists) -------

Write-Section "6. Clearing Outlook resiliency data"

$resiliencyBase = "$hku\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency"

$resiliencyKeys = @("DisabledItems", "CrashingAddinList", "NotificationReminderAddinData")
$resiliencyCleared = $false

foreach ($key in $resiliencyKeys) {
    $fullPath = "$resiliencyBase\$key"
    if (Test-Path $fullPath) {
        if ($PSCmdlet.ShouldProcess($fullPath, "Clear resiliency data")) {
            Remove-Item $fullPath -Force -ErrorAction SilentlyContinue
            Write-Status "Cleared: $key" "Success"
            $resiliencyCleared = $true
        }
    }
}

# Also set DoNotDisableAddinList as a belt-and-braces backup
$doNotDisablePath = "$resiliencyBase\DoNotDisableAddinList"
if ($PSCmdlet.ShouldProcess($doNotDisablePath, "Set DoNotDisableAddinList")) {
    New-Item -Path $doNotDisablePath -Force -ErrorAction SilentlyContinue | Out-Null
    foreach ($addin in $AddinsToProtect) {
        Set-ItemProperty $doNotDisablePath -Name $addin -Value 1 -Type DWord -ErrorAction SilentlyContinue
    }
    Write-Status "DoNotDisableAddinList set as backup." "Success"
}

if (-not $resiliencyCleared) {
    Write-Status "No resiliency data to clear." "Skip"
}

# --- 7. Word/Office startup optimisations -----------------------------------

Write-Section "7. Applying Office startup optimisations"

# Disable Word start screen
$wordOptsPath = "$hku\SOFTWARE\Microsoft\Office\16.0\Word\Options"
if (Test-Path $wordOptsPath) {
    if ($PSCmdlet.ShouldProcess($wordOptsPath, "Disable Word start screen")) {
        Set-ItemProperty $wordOptsPath -Name "DisableBootToOfficeStart" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Write-Status "Disabled Word start screen." "Success"
    }
} else {
    Write-Status "Word options key not found - Word may not have been opened yet." "Skip"
}

# Disable Excel start screen
$excelOptsPath = "$hku\SOFTWARE\Microsoft\Office\16.0\Excel\Options"
if (Test-Path $excelOptsPath) {
    if ($PSCmdlet.ShouldProcess($excelOptsPath, "Disable Excel start screen")) {
        Set-ItemProperty $excelOptsPath -Name "DisableBootToOfficeStart" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        Write-Status "Disabled Excel start screen." "Success"
    }
}

# Disable Office animations
$graphicsPath = "$hku\SOFTWARE\Microsoft\Office\16.0\Common\Graphics"
if ($PSCmdlet.ShouldProcess($graphicsPath, "Disable Office animations")) {
    New-Item -Path $graphicsPath -Force -ErrorAction SilentlyContinue | Out-Null
    Set-ItemProperty $graphicsPath -Name "DisableAnimations" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
    Write-Status "Disabled Office animations." "Success"
}

# Fix window focus - stop splash screens appearing behind other windows
$desktopPath = "$hku\Control Panel\Desktop"
if (Test-Path $desktopPath) {
    if ($PSCmdlet.ShouldProcess($desktopPath, "Fix foreground window focus")) {
        Set-ItemProperty $desktopPath -Name "ForegroundLockTimeout" -Value 0 -Type DWord -ErrorAction SilentlyContinue
        Write-Status "Fixed foreground window focus (ForegroundLockTimeout=0)." "Success"
    }
}

# --- Summary -----------------------------------------------------------------

Write-Host ""
Write-Host "=== CLEANUP COMPLETE ===" -ForegroundColor Green
Write-Host ""
Write-Host "  What was done:" -ForegroundColor White
Write-Host "  - Removed dead Outlook add-ins (iManage, ICQ, Lync, etc.)" -ForegroundColor Cyan
Write-Host "  - Cleared .NET assembly cache (dl3)" -ForegroundColor Cyan
Write-Host "  - Checked/disabled Fusion logging" -ForegroundColor Cyan
Write-Host "  - Cleaned iManage registry remnants" -ForegroundColor Cyan
Write-Host "  - Protected required Outlook add-ins from auto-disable" -ForegroundColor Cyan
Write-Host "  - Cleared Outlook resiliency (disabled add-in) data" -ForegroundColor Cyan
Write-Host "  - Applied Office startup optimisations" -ForegroundColor Cyan
Write-Host ""
Write-Host "  IMPORTANT - First launch after cleanup:" -ForegroundColor Yellow
Write-Host "  The first time each Office app opens it will be a bit slower" -ForegroundColor Yellow
Write-Host "  while the assembly cache rebuilds. This is normal and one-time." -ForegroundColor Yellow
Write-Host "  Second launch onwards will be fast." -ForegroundColor Yellow
Write-Host ""
Write-Host "  If user needs to log off/on for foreground fix to take effect." -ForegroundColor Yellow
Write-Host ""
