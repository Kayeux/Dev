<#
.SYNOPSIS
    Simplified "SCCM Killer" script with improved checks to avoid errors on devices without SCCM.
.DESCRIPTION
    - Checks if SCCM is actually installed before uninstall steps.
    - If SCCM is installed, silently uninstalls, removes leftover directories & registry keys, then forces an Intune sync.
    - If not installed, logs a message and exits gracefully.
#>

[CmdletBinding()]
param()

###############################################################################
# 0. Logging Function
###############################################################################
function Write-Log {
    param ($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "$timestamp : $Message"
}

###############################################################################
# 1. Detection: Is SCCM Installed?
###############################################################################
# We'll look for:
#   1) The ccmsetup.exe path
#   2) A known SCCM registry key (e.g. HKLM:\SOFTWARE\Microsoft\CCM)
# If either is found, we assume SCCM is installed.

$SccmRegistryKey = "HKLM:\SOFTWARE\Microsoft\CCM"
$sccmPath        = "$env:WinDir\ccmsetup\ccmsetup.exe"
$SccmInstalled   = $false

if (Test-Path $SccmRegistryKey -PathType Container) {
    $SccmInstalled = $true
}
elseif (Test-Path $sccmPath) {
    $SccmInstalled = $true
}

if (-not $SccmInstalled) {
    Write-Log "SCCM not detected (no known registry keys or ccmsetup.exe). Exiting script."
    return
}

Write-Log "SCCM detected. Proceeding with removal..."

###############################################################################
# 2. Uninstall SCCM
###############################################################################
if (Test-Path $sccmPath) {
    Write-Log "Initiating silent SCCM uninstall..."
    Start-Process -FilePath $sccmPath -ArgumentList "/uninstall /silent" -Wait
    Write-Log "SCCM client uninstall completed."
}
else {
    # It's possible the registry key is there but the ccmsetup.exe is missing.
    Write-Log "ccmsetup.exe not found, but registry key present. Attempting residual cleanup."
}

###############################################################################
# 3. Cleanup Directories
###############################################################################
$folders = @(
    "$env:WinDir\CCM",
    "$env:WinDir\ccmsetup",
    "$env:WinDir\ccmcache"
)

foreach ($folder in $folders) {
    if (Test-Path $folder) {
        Write-Log "Removing folder: $folder"
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Log "Folder already removed or not found: $folder"
    }
}

###############################################################################
# 4. Cleanup Registry (Do this last to confirm SCCM was installed)
###############################################################################
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\CCM",
    "HKLM:\SOFTWARE\Microsoft\CCMSetup",
    "HKLM:\SOFTWARE\Microsoft\SMS"
)

foreach ($reg in $registryPaths) {
    if (Test-Path $reg) {
        Write-Log "Removing registry key: $reg"
        Remove-Item -Path $reg -Recurse -Force -ErrorAction SilentlyContinue
    }
    else {
        Write-Log "Registry key not found or already removed: $reg"
    }
}

###############################################################################
# 5. Force Intune Sync (Optional)
###############################################################################
Write-Log "Triggering Intune Sync..."
try {
    Get-ScheduledTask -TaskName "PushLaunch" | Start-ScheduledTask
    Write-Log "Intune Sync command executed."
}
catch {
    Write-Log "Could not start PushLaunch task. Possibly not present."
}

Write-Log "SCCM Removal process complete. A reboot is recommended."