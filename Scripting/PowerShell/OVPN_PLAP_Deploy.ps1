<#
.SYNOPSIS
  AIO script to:
   1) Locate an .ovpn file in two possible paths (recursively)
   2) Import PLAP registry file
   3) Insert management lines for SBL/PLAP
   4) Move it to config-auto
   5) Restart the OpenVPN service

.DESCRIPTION
  Searches under:
    1) C:\Users\$env:USERNAME\OpenVPN\config\   (recursively)
    2) C:\Program Files\OpenVPN\config\         (recursively)
  If multiple .ovpn files are found in a path, picks the first one.

.PARAMETER PlapRegFile
  Path to openvpn-plap-install.reg

.PARAMETER DestOvpnName
  The name to use for the final .ovpn file in config-auto

.PARAMETER NeedTOTP
  Add 'auth-retry interact' if TOTP-based MFA is used
#>

param (
  [string]$PlapRegFile   = "C:\Program Files\OpenVPN\bin\openvpn-plap-install.reg",
  [string]$DestOvpnName  = "MyVPN-PLAP.ovpn",
  [switch]$NeedTOTP
)

# --- CONFIGURABLE PATHS ---
$OpenVPNInstallDir = "C:\Program Files\OpenVPN"
$ConfigAutoDir     = Join-Path $OpenVPNInstallDir "config-auto"
$OpenVPNService    = "OpenVPNService"

# Potential source directories for the .ovpn file (searching recursively)
$OvpnPathsToCheck = @(
  "C:\Users\$($env:USERNAME)\OpenVPN\config\",
  "C:\Program Files\OpenVPN\config\"
)

# Required lines for pre-logon (PLAP/SBL)
$RequiredManagementLines = @(
  "management 127.0.0.1 12345",
  "management-hold",
  "management-query-passwords"
)
if ($NeedTOTP) {
  $RequiredManagementLines += "auth-retry interact"
}

Write-Host "`n===== 1. Checking Admin Privileges ====="
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "ERROR: This script must be run as Administrator. Exiting."
    return
}
else {
    Write-Host "Running as Administrator."
}

Write-Host "`n===== 2. Locating .ovpn File in Known Paths (Recursive) ====="
$foundOvpn = $null

foreach ($path in $OvpnPathsToCheck) {
    if (Test-Path $path) {
        # Recursively find .ovpn files
        $ovpnFiles = Get-ChildItem -Path $path -Recurse -Filter *.ovpn -File -ErrorAction SilentlyContinue
        if ($ovpnFiles) {
            if ($ovpnFiles.Count -gt 1) {
                Write-Warning "Multiple .ovpn files found in '$path'. Using the first one: $($ovpnFiles[0].Name)"
            }
            $foundOvpn = $ovpnFiles[0].FullName
            Write-Host "Found .ovpn file: $foundOvpn"
            break
        }
        else {
            Write-Host "No .ovpn files found in '$path' (recursively). Checking next..."
        }
    }
    else {
        Write-Host "Path does not exist: $path"
    }
}

if (-Not $foundOvpn) {
    Write-Error "No .ovpn file found in either user config or Program Files config paths. Exiting."
    return
}

Write-Host "`n===== 3. Importing PLAP Registry File ====="
if (Test-Path $PlapRegFile) {
    try {
        Write-Host "Importing: $PlapRegFile"
        reg import "$PlapRegFile" | Out-Null
        Write-Host "Successfully imported PLAP registry keys."
    }
    catch {
        Write-Warning "Failed to import $PlapRegFile : $($_.Exception.Message)"
    }
}
else {
    Write-Warning "PLAP .reg file not found at: $PlapRegFile. Make sure the path is correct."
}

Write-Host "`n===== 4. Modifying OVPN for SBL (Adding management lines) ====="
$fileContent = Get-Content -Path $foundOvpn -ErrorAction Stop

foreach ($line in $RequiredManagementLines) {
    if ($fileContent -notcontains $line) {
        Write-Host "Adding line: $line"
        $fileContent += $line
    }
}

# Construct final path in config-auto
$DestinationFile = Join-Path $ConfigAutoDir $DestOvpnName

Write-Host "Copying modified OVPN to $DestinationFile"
try {
    if (-Not (Test-Path $ConfigAutoDir)) {
        New-Item -Path $ConfigAutoDir -ItemType Directory -Force | Out-Null
    }
    $fileContent | Out-File -FilePath $DestinationFile -Encoding ASCII
    Write-Host "Modified .ovpn successfully placed in config-auto."
}
catch {
    Write-Warning "Failed to write file to $DestinationFile : $($_.Exception.Message)"
    return
}

Write-Host "`n===== 5. Restarting OpenVPN Service ====="
try {
    Write-Host "Stopping $OpenVPNService..."
    Stop-Service $OpenVPNService -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    Write-Host "Starting $OpenVPNService..."
    Start-Service $OpenVPNService
    Write-Host "OpenVPN service restarted successfully."
}
catch {
    Write-Warning "Failed to restart $OpenVPNService : $($_.Exception.Message)"
}

Write-Host "`n===== ALL DONE ====="
Write-Host "Pre-logon provider is installed, OVPN config is updated, and the service is running."
Write-Host "Lock or reboot to see the VPN icon at the Windows sign-in screen."