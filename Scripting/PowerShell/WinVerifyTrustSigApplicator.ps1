# PowerShell Script to Apply WinVerifyTrust Signature Verification Patch

# Function to check if a registry key exists
function Test-RegistryKey {
    param (
        [string]$Path,
        [string]$Name
    )
    return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue) -ne $null
}

# Function to set the EnableCertPaddingCheck registry entry
function Set-EnableCertPaddingCheck {
    param (
        [string]$Path,
        [string]$Name,
        [string]$Value
    )
    try {
        # Ensure the registry path exists
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        # Set the registry value
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
        Write-Host "Successfully set $Name in $Path"
    }
    catch {
        Write-Error "Failed to set $Name in $Path. $_"
    }
}

# Function to remove the EnableCertPaddingCheck registry entry
function Remove-EnableCertPaddingCheck {
    param (
        [string]$Path,
        [string]$Name
    )
    try {
        Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        Write-Host "Successfully removed $Name from $Path"
    }
    catch {
        Write-Error "Failed to remove $Name from $Path. $_"
    }
}

# Main Script Execution Starts Here

# Ensure the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Warning "You do not have Administrator rights to run this script. Please run PowerShell as an Administrator and try again."
    exit
}

# Define registry paths and value name
$baseRegistryPath = "HKLM:\Software\Microsoft\Cryptography\Wintrust\Config"
$valueName = "EnableCertPaddingCheck"
$valueData = "1"

# Determine if the operating system is 64-bit
$is64BitOS = [Environment]::Is64BitOperatingSystem

# Initialize a list to hold registry paths to check/set
$registryPaths = @()

if ($is64BitOS) {
    # For 64-bit OS, set both 64-bit and 32-bit registry paths
    $registryPaths += $baseRegistryPath
    $registryPaths += "HKLM:\Software\Wow6432Node\Microsoft\Cryptography\Wintrust\Config"
}
else {
    # For 32-bit OS, set only the base registry path
    $registryPaths += $baseRegistryPath
}

# Flag to determine if patch is already applied
$patchAlreadyExists = $true

# Check if the registry entries already exist
foreach ($path in $registryPaths) {
    if (-not (Test-RegistryKey -Path $path -Name $valueName)) {
        $patchAlreadyExists = $false
        break
    }
}

if ($patchAlreadyExists) {
    Write-Host "Patch Already Exists."
    exit
}

# Apply the patch by setting the registry entries
foreach ($path in $registryPaths) {
    Set-EnableCertPaddingCheck -Path $path -Name $valueName -Value $valueData
}

# Verify that the registry entries have been set
$patchAppliedSuccessfully = $true

foreach ($path in $registryPaths) {
    if (-not (Test-RegistryKey -Path $path -Name $valueName)) {
        Write-Error "Failed to verify the registry entry in $path."
        $patchAppliedSuccessfully = $false
    }
}

if ($patchAppliedSuccessfully) {
    Write-Host "Patch Applied Successfully. A system reboot is pending for the changes to take effect."
}
else {
    Write-Error "Patch was not fully applied. Please check the error messages above."
}
