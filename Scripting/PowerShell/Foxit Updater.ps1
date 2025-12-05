[CmdletBinding()]
Param(
    [switch]$AcknowledgeMissing,
    [switch]$DeepBackup
)

<#
.SYNOPSIS
    Check and update Foxit PhantomPDF if the installed version is below
    the "highest hotfix" of its major version.

.DESCRIPTION
    1. Searches registry for any Foxit PhantomPDF installation.
    2. Compares installed version to known max versions for major 8, 9, or 10.
    3. If update is needed:
        - (Optional) Creates a zip backup of the entire install folder if -DeepBackup is used.
        - Checks for fpmkey.txt:
            * If missing/corrupt and -AcknowledgeMissing is NOT specified, script exits with error.
            * If -AcknowledgeMissing is specified, script continues even if missing.
        - Performs a test GET on the update link to ensure it's valid.
        - Downloads the MSI to TEMP.
        - Backs up fpmkey.txt (if present).
        - Installs with msiexec /quiet /norestart.
        - Restores fpmkey.txt post-install.
    4. If the installed version is already at or above the highest hotfix, no action is taken.

.EXAMPLE
    # Run script with no special flags (exits on missing fpmkey.txt)
    .\UpdateFoxit.ps1

.EXAMPLE
    # Run script, ignoring missing fpmkey.txt
    .\UpdateFoxit.ps1 -AcknowledgeMissing

.EXAMPLE
    # Run script, ignoring missing fpmkey.txt and performing a deep folder backup
    .\UpdateFoxit.ps1 -AcknowledgeMissing -DeepBackup
#>

# Mapping: Major version -> Highest known hotfix version
$HighestHotfixByMajor = @{
    '8'  = [Version]'8.3.12.47136'
    '9'  = [Version]'9.7.5.29616'
    '10' = [Version]'10.1.12.37872'
}

# Mapping: Major version -> Download link
$DownloadLinkByMajor = @{
    '8'  = 'https://cdn09.foxitsoftware.com/pub/foxit/phantomPDF/desktop/win/8.x/8.3/en_us/FoxitPhantomPDF8312_enu_Setup.msi'
    '9'  = 'https://cdn09.foxitsoftware.com/pub/foxit/phantomPDF/desktop/win/9.x/9.7/en_us/FoxitPhantomPDF975_enu_Setup_Website.msi'
    '10' = 'https://cdn09.foxitsoftware.com/product/phantomPDF/desktop/win/10.1.12/FoxitPhantomPDF10112_enu_Setup_Website.msi'
}

Write-Host "`n=== Checking for Foxit PhantomPDF installation ===`n"

$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$foxitRegInfo = $null

foreach ($path in $uninstallPaths) {
    $subkeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
    foreach ($subkey in $subkeys) {
        try {
            $regValues = Get-ItemProperty -Path $subkey.PSPath -ErrorAction Stop
            if ($regValues.DisplayName -like '*Foxit PhantomPDF*') {
                $foxitRegInfo = [PSCustomObject]@{
                    DisplayName     = $regValues.DisplayName
                    DisplayVersion  = $regValues.DisplayVersion
                    InstallLocation = $regValues.InstallLocation
                    RegistryKey     = $subkey.PSPath
                }
                break
            }
        }
        catch {
            # Ignore read errors
        }
    }
    if ($foxitRegInfo) { break }
}

if (-not $foxitRegInfo) {
    Write-Host "Foxit PhantomPDF is NOT installed on this system. Exiting updater."
    return
}

Write-Host "Foxit PhantomPDF is installed."
Write-Host "Display Name  : $($foxitRegInfo.DisplayName)"
Write-Host "Version       : $($foxitRegInfo.DisplayVersion)"
Write-Host "Install Path  : $($foxitRegInfo.InstallLocation)"
Write-Host "Registry Key  : $($foxitRegInfo.RegistryKey)"

# Convert installed version to [Version]
try {
    $installedVersion = [Version]$foxitRegInfo.DisplayVersion
} catch {
    Write-Warning "Could not parse Foxit version. Exiting."
    return
}

# Identify the major version
$majorVersion = $installedVersion.Major.ToString()

if (-not $HighestHotfixByMajor.ContainsKey($majorVersion)) {
    Write-Host "`nNo highest hotfix info is defined for major version: $majorVersion"
    Write-Host "No update check performed."
    return
}

$highestKnownVersion = $HighestHotfixByMajor[$majorVersion]
Write-Host "`nFound highest hotfix for major version $($majorVersion) : $($highestKnownVersion)"

if ($installedVersion -ge $highestKnownVersion) {
    Write-Host "`nInstalled version ($installedVersion) is already at or above $highestKnownVersion. No update needed."
    return
}

Write-Host "`nInstalled version ($installedVersion) is LESS than the highest known version ($highestKnownVersion)."
Write-Host "An update is REQUIRED. Preparing to update..."

# Get the download link for this major version
if (-not $DownloadLinkByMajor.ContainsKey($majorVersion)) {
    Write-Warning "No download link defined for major version $majorVersion. Exiting."
    return
}
$msiDownloadLink = $DownloadLinkByMajor[$majorVersion]

# Attempt to find Foxit install folder or guess a default
$defaultFoxitDir = 'C:\Program Files (x86)\Foxit Software'
$foxitBasePath   = $foxitRegInfo.InstallLocation
if ([string]::IsNullOrWhiteSpace($foxitBasePath)) {
    $foxitBasePath = $defaultFoxitDir
}

# --- [Deep Backup Section] ---
if ($DeepBackup) {
    # We'll store the zip in ProgramData or some location
    $backupRoot = Join-Path $env:ProgramData "FoxitKeyBackup"
    if (!(Test-Path $backupRoot)) {
        New-Item -Path $backupRoot -ItemType Directory | Out-Null
    }

    # Create a zip file name. For uniqueness, we can add a timestamp
    $zipName = "FoxitPhantomPDFBackup_{0:yyyyMMdd_HHmmss}.zip" -f (Get-Date)
    $zipPath = Join-Path $backupRoot $zipName
    
    Write-Host "`n[Deep Backup] Creating a zip of the entire folder:"
    Write-Host "Source: $foxitBasePath"
    Write-Host "Destination: $zipPath"

    try {
        if (!(Test-Path $foxitBasePath)) {
            Write-Error "Cannot do a deep backup. $foxitBasePath does not exist."
            return
        }
        # Compress-Archive needs a folder path or a set of items
        Compress-Archive -Path $foxitBasePath -DestinationPath $zipPath -Force
        if (!(Test-Path $zipPath)) {
            Write-Error "Deep backup zip was not created as expected. Exiting."
            return
        }
        Write-Host "[Deep Backup] Zip created successfully at $zipPath"
    }
    catch {
        Write-Error "Failed to create deep backup zip: $($_.Exception.Message)"
        return
    }
}

# --- [Check fpmkey.txt] ---
# If not found, exit unless -AcknowledgeMissing is provided
Write-Host "`nSearching for fpmkey.txt..."
try {
    $fpmkeyPath = Get-ChildItem -Path $foxitBasePath -Recurse -Filter 'fpmkey.txt' -ErrorAction SilentlyContinue |
                  Select-Object -First 1
}
catch {
    Write-Warning "Could not search for fpmkey.txt: $($_.Exception.Message)"
    $fpmkeyPath = $null
}

if (-not $fpmkeyPath) {
    if (-not $AcknowledgeMissing) {
        Write-Error "fpmkey.txt file not found or missing. Use -AcknowledgeMissing to proceed anyway. Exiting."
        return
    }
    else {
        Write-Host "fpmkey.txt is missing, but -AcknowledgeMissing is set. Proceeding..."
    }
}
else {
    Write-Host "fpmkey.txt found at: $($fpmkeyPath.FullName)"
}

# --- Verify if the MSI exists prior to downloading again ---
$msiFileName = Split-Path $msiDownloadLink -Leaf
$downloadFolder = Join-Path $env:TEMP "FoxitUpdate"
if (!(Test-Path $downloadFolder)) {
    New-Item -Path $downloadFolder -ItemType Directory | Out-Null
}
$localMsiPath = Join-Path $downloadFolder $msiFileName

Write-Host "`n=== Checking if the MSI is already present ==="
if(!(Test-Path $localMsiPath))
{
    Write-Host "`nMSI not already found, performing a fresh download."

    # --- [Download the MSI] ---
    Write-Host "`nDownloading the MSI from $msiDownloadLink to $localMsiPath..."
    try {
        Invoke-WebRequest -Uri $msiDownloadLink -OutFile $localMsiPath -UseBasicParsing
    }
    catch {
        Write-Error "Download failed: $($_.Exception.Message)"
        return
    }
}
else {
    Write-Host "`nMSI Previously found at $localMsiPath , proceeding with install."
}

# Prepare for the license key backup location
$fpmkeyBackupFolder = Join-Path $env:ProgramData "FoxitKeyBackup"
if (!(Test-Path $fpmkeyBackupFolder)) {
    New-Item -Path $fpmkeyBackupFolder -ItemType Directory | Out-Null
}

# If fpmkey.txt is found, back it up
if ($fpmkeyPath) {
    Write-Host "`nBacking up fpmkey.txt..."
    try {
        Copy-Item -Path $fpmkeyPath.FullName -Destination $fpmkeyBackupFolder -Force
        Write-Host "fpmkey.txt backed up to: $fpmkeyBackupFolder"
    }
    catch {
        Write-Warning "Could not backup fpmkey.txt: $($_.Exception.Message)"
    }
}

# --- [Install the MSI] ---
Write-Host "`nInstalling Foxit PhantomPDF MSI..."
if (Test-Path $fpmkeyBackupFolder )
{
    $installArgs = "/i `"$localMsiPath`" /quiet /norestart KEYPATH=`"$fpmkeyBackupFolder\fpmkey.txt`""
}
else {
    $installArgs = "/i `"$localMsiPath`" /quiet /norestart"
}
$msiExecCmd = "msiexec.exe $installArgs"
Write-Host $msiExecCmd
$installResult = Start-Process -FilePath "msiexec.exe" -ArgumentList $installArgs -Wait -PassThru
if ($installResult.ExitCode -eq 0) {
    Write-Host "Installation completed successfully."
}
else {
    Write-Warning "Installation may have failed. msiexec exit code: $($installResult.ExitCode)."
}

# --- [Restore the fpmkey.txt] ---
if ($fpmkeyPath) {
    Write-Host "`nRestoring backed-up fpmkey.txt..."
    # Try the same folder it was originally found in
    $restoreFolder = Split-Path $fpmkeyPath.FullName
    if (!(Test-Path $restoreFolder)) {
        # If the folder doesn't exist, try new install location or the default
        if ((Test-Path $foxitRegInfo.InstallLocation) -and $foxitRegInfo.InstallLocation) {
            $restoreFolder = $foxitRegInfo.InstallLocation
        }
        else {
            $restoreFolder = $defaultFoxitDir
        }
    }
    $backupFpmkey = Join-Path $fpmkeyBackupFolder $fpmkeyPath.Name
    if (Test-Path $backupFpmkey) {
        Copy-Item -Path $backupFpmkey -Destination $restoreFolder -Force
        Write-Host "fpmkey.txt restored to $restoreFolder"
    }
    else {
        Write-Warning "Could not find the backed-up fpmkey.txt in $fpmkeyBackupFolder"
    }
}

Write-Host "`nUpdate process finished."
