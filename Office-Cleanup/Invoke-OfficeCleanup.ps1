#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Wallace Office Startup Cleanup Script
    Fixes common causes of slow Word, Excel, and Outlook startup.

.DESCRIPTION
    This script performs the following cleanup tasks:
    1. Removes dead/unwanted Outlook COM add-ins (iManage, ICQ, Lync, etc.)
    2. Clears the .NET assembly download cache (dl3) which accumulates stale DLLs
    3. Disables .NET Fusion assembly binding logging if enabled (massive perf killer)
    4. Cleans up iManage/Interwoven registry remnants
    5. Prevents Outlook from auto-disabling required add-ins (Workshare, Litera, etc.)

.PARAMETER UserName
    The username whose profile to clean. Defaults to the currently logged-in user.
    Use this when running as a different admin account.

.PARAMETER WhatIf
    Shows what would be changed without making any changes.

.EXAMPLE
    .\Invoke-OfficeCleanup.ps1
    Runs cleanup for the current user.

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
# Add-ins to REMOVE (dead/legacy/unwanted)
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

# Add-ins that Outlook must NOT auto-disable (resiliency protection)
$AddinsToProtect = @(
    "Workshare.OutlookRibbon.Addin",
    "zzzDocsCorp.pdfDocs.OutlookAddIn",
    "MetaCompliance.Reporter.Email",
    "MimecastServicesForOutlook.Connect",
    "zzz.SafeSend",
    "TeamsAddin.FastConnect"
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

# --- Resolve target user profile path ----------------------------------------

if ($UserName) {
    $userProfile = "C:\Users\$UserName"
} else {
    # Get the logged-in user (not the admin running the script)
    $loggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
    if ($loggedInUser -and $loggedInUser -match "\\(.+)$") {
        $resolvedUser = $Matches[1]
        $userProfile = "C:\Users\$resolvedUser"
    } else {
        $userProfile = $env:USERPROFILE
        $resolvedUser = $env:USERNAME
    }
}

if (-not (Test-Path $userProfile)) {
    Write-Host "ERROR: User profile not found at $userProfile" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Wallace Office Startup Cleanup" -ForegroundColor White
Write-Host "Target user profile: $userProfile" -ForegroundColor Cyan
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
    # Also check HKCU
    $hkcuPath = "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins\$addin"
    if (Test-Path $hkcuPath) {
        if ($PSCmdlet.ShouldProcess($hkcuPath, "Remove add-in")) {
            Remove-Item -Path $hkcuPath -Force -ErrorAction SilentlyContinue
            Write-Status "Removed: $addin (from HKCU)" "Success"
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
            Write-Status "NOTE: First app launch will be slower as cache rebuilds. Second launch will be fast." "Warning"
        } else {
            Write-Status "Partial clear - some files were locked. A reboot may be needed." "Warning"
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

# Also clean up Fusion log output directory if it exists
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
    "HKCU:\SOFTWARE\Interwoven",
    "HKLM:\SOFTWARE\iManage",
    "HKLM:\SOFTWARE\WOW6432Node\iManage",
    "HKCU:\SOFTWARE\iManage"
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

# --- 5. Protect required add-ins from Outlook resiliency --------------------

Write-Section "5. Protecting required add-ins from auto-disable"

# Load the target user's HKCU hive if running as a different admin
$useHKU = $false
$hkuPath = ""
if ($UserName -and $UserName -ne $env:USERNAME) {
    # Try to find the user's SID
    try {
        $userSID = (New-Object System.Security.Principal.NTAccount($UserName)).Translate(
            [System.Security.Principal.SecurityIdentifier]).Value
        $hkuPath = "Registry::HKU\$userSID"
        if (Test-Path $hkuPath) {
            $useHKU = $true
        } else {
            # Need to load the hive
            $ntUserDat = Join-Path $userProfile "NTUSER.DAT"
            if (Test-Path $ntUserDat) {
                $tempKey = "HKU\TempCleanup_$UserName"
                reg load $tempKey $ntUserDat 2>$null
                $hkuPath = "Registry::$tempKey"
                $useHKU = $true
                Write-Status "Loaded user registry hive for $UserName" "Info"
            }
        }
    } catch {
        Write-Status "Cannot access $UserName's HKCU - resiliency fix will apply to current admin only" "Warning"
    }
}

$resiliencyBase = if ($useHKU) { "$hkuPath\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency" }
                  else { "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency" }

$doNotDisablePath = "$resiliencyBase\DoNotDisableAddinList"

if ($PSCmdlet.ShouldProcess($doNotDisablePath, "Set add-in protection")) {
    New-Item -Path $doNotDisablePath -Force -ErrorAction SilentlyContinue | Out-Null
    foreach ($addin in $AddinsToProtect) {
        Set-ItemProperty $doNotDisablePath -Name $addin -Value 1 -Type DWord -ErrorAction SilentlyContinue
        Write-Status "Protected: $addin" "Success"
    }

    # Clear any existing disabled items list
    $disabledItemsPath = "$resiliencyBase\DisabledItems"
    if (Test-Path $disabledItemsPath) {
        Remove-Item $disabledItemsPath -Force -ErrorAction SilentlyContinue
        Write-Status "Cleared disabled items list." "Success"
    }
}

# Unload temp hive if we loaded one
if ($useHKU -and $hkuPath -match "TempCleanup") {
    $tempKey = $hkuPath -replace "Registry::", ""
    [gc]::Collect()
    reg unload $tempKey 2>$null
}

# --- Summary -----------------------------------------------------------------

Write-Host ""
Write-Host "=== CLEANUP COMPLETE ===" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "  1. Open each Office app once to rebuild the assembly cache." -ForegroundColor Cyan
Write-Host "     First launch will be a bit slower, second launch will be fast." -ForegroundColor Cyan
Write-Host "  2. Open Outlook and verify add-ins are correct:" -ForegroundColor Cyan
Write-Host "     File > Options > Add-ins > COM Add-ins > Go" -ForegroundColor Cyan
Write-Host ""
