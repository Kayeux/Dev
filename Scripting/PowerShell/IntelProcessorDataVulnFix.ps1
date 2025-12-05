<#
.SYNOPSIS
    Mitigation Script for Speculative Execution Vulnerabilities

.DESCRIPTION
    This script applies various mitigations for speculative execution vulnerabilities based on processor type and user-specified parameters.
    It now automatically applies the Mitigate0001 patch for Intel processors when any mitigation is applied.
    An additional parameter allows you to choose whether to disable Hyper-Threading when using MitigateAll.
    The script outputs the mitigation status to a Ninja custom field titled MAPatchState.

.PARAMETER CheckStatus
    Default parameter. Checks the existence and values of the registry keys without making changes.

.PARAMETER MitigateSSB
    Applies mitigations for Speculative Store Bypass (SSB). The mitigation value differs based on processor type.

.PARAMETER MitigateMMIO
    Applies mitigations for MMIO vulnerabilities.

.PARAMETER MitigateSpectre2
    Applies mitigations for Spectre Variant 2.

.PARAMETER MitigateAll
    Applies all mitigations. Use -DisableHyperThreading to disable Hyper-Threading.

.PARAMETER DisableHyperThreading
    When used with -MitigateAll, disables Hyper-Threading.

.PARAMETER DisableMitigations
    Disables all mitigations by setting /d = 3.

.PARAMETER Force
    Forces the script to modify existing registry keys even if they already exist.

.EXAMPLE
    .\MitigationScript.ps1 -MitigateSSB
    .\MitigationScript.ps1 -MitigateAll -DisableHyperThreading -Force

.NOTES
    This script requires Administrator privileges to modify registry keys.
    A reboot is required for changes to take effect.
#>

# Parameters
[CmdletBinding()]
Param(
    [Parameter(Mandatory = $false)]
    [Switch]$CheckStatus,

    [Parameter(Mandatory = $false)]
    [Switch]$MitigateSSB,

    [Parameter(Mandatory = $false)]
    [Switch]$MitigateMMIO,

    [Parameter(Mandatory = $false)]
    [Switch]$MitigateSpectre2,

    [Parameter(Mandatory = $false)]
    [Switch]$MitigateAll,

    [Parameter(Mandatory = $false)]
    [Switch]$DisableHyperThreading,

    [Parameter(Mandatory = $false)]
    [Switch]$DisableMitigations,

    [Parameter(Mandatory = $false)]
    [Switch]$Force
)

# Default action is to check status if no parameters are provided
If (-Not ($PSBoundParameters.Keys | Where-Object {$_ -ne 'Force'})) {
    $CheckStatus = $true
}

# Define registry paths and keys
$RegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
$FeatureSettingsOverride = "FeatureSettingsOverride"
$FeatureSettingsOverrideMask = "FeatureSettingsOverrideMask"
$HyperVRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization"

# Define backup path with date
$DateStamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
$BackupPath = "$env:SystemRoot\RegistryBackup-$DateStamp.reg"

# Initialize tracking variables
$ExistingPatches = @()
$AppliedPatches = @()
$OverwrittenPatches = @()
$PatchesAlreadyExist = $false

# Function to detect processor type
Function Get-ProcessorType {
    $processor = Get-CimInstance Win32_Processor | Select-Object -First 1
    $manufacturer = $processor.Manufacturer
    If ($manufacturer -match "Intel") {
        Return "Intel"
    } ElseIf ($manufacturer -match "AMD") {
        Return "AMD"
    } Else {
        Return "Unknown"
    }
}

# Function to check if Hyper-V is installed
Function Is-HyperVInstalled {
    $hyperVFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-All -ErrorAction SilentlyContinue
    If ($hyperVFeature -and $hyperVFeature.State -eq 'Enabled') {
        Return $true
    } Else {
        Return $false
    }
}

# Function to back up the registry keys
Function Backup-RegistryKeys {
    # Backs up the current registry settings to a .reg file
    Try {
        reg export "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management" $BackupPath /y
        Write-Host "Registry keys backed up to $BackupPath"
    } Catch {
        Write-Warning "Failed to back up registry keys."
    }
}

# Function to output current registry values
Function Check-RegistryStatus {
    Write-Host "Checking registry keys..."

    Try {
        $currentFSO = Get-ItemPropertyValue -Path $RegPath -Name $FeatureSettingsOverride -ErrorAction Stop
        Write-Host "FeatureSettingsOverride exists with value: 0x$([Convert]::ToString($currentFSO,16).PadLeft(8,'0'))"
    } Catch {
        Write-Host "FeatureSettingsOverride does not exist."
    }

    Try {
        $currentFSOM = Get-ItemPropertyValue -Path $RegPath -Name $FeatureSettingsOverrideMask -ErrorAction Stop
        Write-Host "FeatureSettingsOverrideMask exists with value: 0x$([Convert]::ToString($currentFSOM,16).PadLeft(8,'0'))"
    } Catch {
        Write-Host "FeatureSettingsOverrideMask does not exist."
    }
}

# New line to ease readability
Write-Host ""

# Detect processor type
$processorType = Get-ProcessorType
Write-Host "Processor Type Detected: $processorType"

# Check if Hyper-V is installed
$HyperVInstalled = Is-HyperVInstalled
If ($HyperVInstalled) {
    Write-Host "Hyper-V is installed."
} Else {
    Write-Host "Hyper-V is not installed."
}

# Back up registry keys
Backup-RegistryKeys

# Main logic
If ($CheckStatus) {
    Check-RegistryStatus
    Exit
}

# Stage global variables
$global:forceFlag = $false

# Function to apply mitigations
Function Apply-Mitigation {
    Param(
        [string]$DesiredFSOHexValue,
        [string]$MitigationName,
        [bool]$Additive = $false,
        [bool]$AMDOnly = $false,
        [bool]$IntelOnly = $false
    )

    # Skip if processor type does not match
    If ($AMDOnly -and $processorType -ne "AMD") {
        Write-Host "$MitigationName applies only to AMD processors. Skipping..."
        Return
    }
    If ($IntelOnly -and $processorType -ne "Intel") {
        Write-Host "$MitigationName applies only to Intel processors. Skipping..."
        Return
    }

    # Convert DesiredFSOHexValue to UInt32
    $DesiredFSOValue = [UInt32]::Parse($DesiredFSOHexValue, 'HexNumber')

    # Check current value of FeatureSettingsOverride
    Try {
        $currentFSO = [UInt32](Get-ItemPropertyValue -Path $RegPath -Name $FeatureSettingsOverride -ErrorAction Stop)
        Write-Host "Current FeatureSettingsOverride value: 0x$([Convert]::ToString($currentFSO,16).PadLeft(8,'0'))"
    } Catch {
        Write-Host "FeatureSettingsOverride does not exist. It will be created."
        $currentFSO = [UInt32]0
    }

    # Check current value of FeatureSettingsOverrideMask
    Try {
        $currentFSOM = [UInt32](Get-ItemPropertyValue -Path $RegPath -Name $FeatureSettingsOverrideMask -ErrorAction Stop)
        Write-Host "Current FeatureSettingsOverrideMask value: 0x$([Convert]::ToString($currentFSOM,16).PadLeft(8,'0'))"
    } Catch {
        Write-Host "FeatureSettingsOverrideMask does not exist. It will be created."
        $currentFSOM = [UInt32]0
    }

    # Determine if desired mitigation is already applied
    $isAlreadyApplied = $false
    If ($Additive) {
        If ( $currentFSO -eq $DesiredFSOValue ) {
            $isAlreadyApplied = $true
        }
        $newFSOValue = $currentFSO -bor $DesiredFSOValue
    } Else {
        If ($currentFSO -eq $DesiredFSOValue) {
            $isAlreadyApplied = $true
        }
        $newFSOValue = $DesiredFSOValue
    }

    If ($isAlreadyApplied) {
        Write-Host "$MitigationName mitigation is already applied."
        # Record existing patches
        $PatchesAlreadyExist = $true
        $ExistingPatches += $MitigationName
        Return
    }

    # If not forcing and value exists, output and exit
    If (-Not $Force -and $currentFSO -ne 0 -and -Not $Additive) {
        Write-Host "FeatureSettingsOverride already exists with value: 0x$([Convert]::ToString($currentFSO,16).PadLeft(8,'0'))"
        Write-Host "Use the -Force parameter to override."
        # Record existing patches
        $PatchesAlreadyExist = $true
        $ExistingPatches += $MitigationName
        Return
    }

    # Check if we are overwriting existing settings
    If ($currentFSO -ne 0 -and $currentFSO -ne $newFSOValue) {
        Write-Host "Overwriting existing FeatureSettingsOverride value."
        # Record overwritten patches
        $OverwrittenPatches += $MitigationName
    }

    Write-Host "New FeatureSettingsOverride value: 0x$([Convert]::ToString($newFSOValue,16).PadLeft(8,'0'))"
    Write-Host "New FeatureSettingsOverrideMask value: 0x00000003"

    # Apply the registry keys with proper data types
    Set-ItemProperty -Path $RegPath -Name $FeatureSettingsOverride -Type DWord -Value $newFSOValue
    Set-ItemProperty -Path $RegPath -Name $FeatureSettingsOverrideMask -Type DWord -Value 3

    # Add Hyper-V registry key if necessary
    If ($HyperVInstalled -and ($MitigationName -ne "Disable Mitigations")) {
        New-ItemProperty -Path $HyperVRegPath -Name "MinVmVersionForCpuBasedMitigations" -PropertyType String -Value "1.0" -Force | Out-Null
        Write-Host "Hyper-V mitigation applied."
    }

    Write-Host "$MitigationName mitigation applied."
    Write-Host "A reboot is required for changes to take effect."

    # Record applied patches
    $AppliedPatches += $MitigationName
}

# Process parameters

# If the processor is Intel and any mitigation is being applied, automatically apply Mitigate0001
If (($processorType -eq "Intel") -and (-not $CheckStatus) -and ($MitigateSSB -or $MitigateMMIO -or $MitigateSpectre2 -or $MitigateAll -or $DisableMitigations)) {
    $Mitigate0001 = $true
}

# Apply SSB Mitigation
If ($MitigateSSB) {
    If ($processorType -eq "Intel") {
        Apply-Mitigation -DesiredFSOHexValue "00000008" -MitigationName "SSB (Intel)"
    } ElseIf ($processorType -eq "AMD") {
        Apply-Mitigation -DesiredFSOHexValue "00000048" -MitigationName "SSB (AMD)"
    } Else {
        Write-Host "Unknown processor type. Cannot apply SSB mitigation."
    }
}

# Apply MMIO Mitigation
If ($MitigateMMIO) {
    Apply-Mitigation -DesiredFSOHexValue "00000000" -MitigationName "MMIO"
}

# Apply Spectre Variant 2 Mitigation
If ($MitigateSpectre2) {
    If ($processorType -eq "Intel") {
        Apply-Mitigation -DesiredFSOHexValue "00000000" -MitigationName "Spectre Variant 2 (Intel)"
    } ElseIf ($processorType -eq "AMD") {
        Apply-Mitigation -DesiredFSOHexValue "00000040" -MitigationName "Spectre Variant 2 (AMD)"
    } Else {
        Write-Host "Unknown processor type. Cannot apply Spectre Variant 2 mitigation."
    }
}

# Apply All Mitigations
If ($MitigateAll) {
    If ($DisableHyperThreading) {
        $DesiredFSOHexValue = "00002048" # Mitigates all vulnerabilities and disables Hyper-Threading
    } Else {
        $DesiredFSOHexValue = "00000048" # Mitigates all vulnerabilities but keeps Hyper-Threading enabled
    }
    Apply-Mitigation -DesiredFSOHexValue $DesiredFSOHexValue -MitigationName "All Mitigations"
}

# Apply CVE-2022-0001 Mitigation (Intel Only)
If ($Mitigate0001) {
    Apply-Mitigation -DesiredFSOHexValue "00800000" -MitigationName "CVE-2022-0001" -Additive $true -IntelOnly $true
}

# Disable All Mitigations
If ($DisableMitigations) {
    Apply-Mitigation -DesiredFSOHexValue "00000003" -MitigationName "Disable Mitigations"
}

# After processing all mitigations, output to Ninja Custom Field
If ($PatchesAlreadyExist -and $AppliedPatches.Count -eq 0) {
    # Case 1: Patches found already, but not set new
    $message = "FOUND:`nPatches Already Existing:`n"
    $message += ($ExistingPatches -join "`n")
} ElseIf ($PatchesAlreadyExist -and $AppliedPatches.Count -gt 0) {
    # Case 2: Patches found already, but then set new (overwritten)
    $message = "Overwritten:`nPatches Already Existing:`n"
    $message += ($ExistingPatches -join "`n")
    $message += "`nPatches Applied:`n"
    $message += ($AppliedPatches -join "`n")
} ElseIf (-Not $PatchesAlreadyExist -and $AppliedPatches.Count -gt 0) {
    # Case 3: Not found already, patches applied
    $message = "NEW:`nPatches Applied:`n"
    $message += ($AppliedPatches -join "`n")
} Else {
    # No patches applied or found
    $message = "No patches applied or found."
}

# Output the message
Write-Host $message

# Output to Ninja Custom Field
Ninja-Property-Set MAPatchState $message
