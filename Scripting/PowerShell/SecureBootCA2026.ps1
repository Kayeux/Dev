<#
.SYNOPSIS
  Secure Boot 2026 CA rotation readiness + remediation helper.

.DESCRIPTION
  - Audits Secure Boot enabled state and Secure Boot CA 2023 rollout status.
  - Flags “opted out” devices (HighConfidenceOptOut=1).
  - Optionally triggers IT-managed deployment (AvailableUpdates=0x5944) and runs the handler task.

  Microsoft references:
    - AvailableUpdates=0x5944 triggers the rollout
    - UEFICA2023Status / UEFICA2023Error are primary monitoring keys
    - Task: \Microsoft\Windows\PI\Secure-Boot-Update

.PARAMETER Remediate
  Applies safe registry changes (clears opt-out + triggers IT-managed update) when eligible.
#>

[CmdletBinding()]
param(
  [switch]$Remediate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-RegDword([string]$Path, [string]$Name) {
  try {
    $v = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    if ($null -eq $v) { return $null }
    return [uint32]$v
  } catch { return $null }
}

function Get-RegString([string]$Path, [string]$Name) {
  try {
    $v = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    if ($null -eq $v) { return $null }
    return [string]$v
  } catch { return $null }
}

function Set-RegDword([string]$Path, [string]$Name, [uint32]$Value) {
  if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
  New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
}

function Remove-RegValue([string]$Path, [string]$Name) {
  try { Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop | Out-Null } catch { }
}

function Get-LatestEvent([int]$Id) {
  try {
    Get-WinEvent -FilterHashtable @{ LogName="System"; Id=$Id } -MaxEvents 1 -ErrorAction Stop
  } catch { return $null }
}

$sbRoot  = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot"
$sbState = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State"
$sbSvc   = "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
$task    = "\Microsoft\Windows\PI\Secure-Boot-Update"

# Secure Boot enabled (registry + API best-effort)
$uefiSecureBootEnabled = Get-RegDword $sbState "UEFISecureBootEnabled"
$confirmSecureBoot = $null
try { $confirmSecureBoot = [bool](Confirm-SecureBootUEFI) } catch { }

# Servicing status
$uefiStatus = Get-RegString $sbSvc "UEFICA2023Status"    # NotStarted / InProgress / Updated
$uefiError  = Get-RegDword  $sbSvc "UEFICA2023Error"

# Assists / triggers
$highConfidenceOptOut = Get-RegDword $sbRoot "HighConfidenceOptOut"
$msManagedOptIn       = Get-RegDword $sbRoot "MicrosoftUpdateManagedOptIn"
$availableUpdates     = Get-RegDword $sbRoot "AvailableUpdates"

# Event checks
$ev1808 = Get-LatestEvent 1808
$ev1801 = Get-LatestEvent 1801

# Determine “updated”
$updated = $false
if ($uefiStatus -eq "Updated") { $updated = $true }
elseif ($null -ne $ev1808)     { $updated = $true }

$denied = @()
if ($uefiSecureBootEnabled -eq 0 -or $confirmSecureBoot -eq $false) { $denied += "SecureBootDisabled" }

$recommended = @()
if ($highConfidenceOptOut -eq 1) { $recommended += "Clear HighConfidenceOptOut (opt back in)" }
if (-not $updated -and ($denied -notcontains "SecureBootDisabled")) { $recommended += "Trigger IT-managed rollout: AvailableUpdates=0x5944" }

$actions = @()
if ($Remediate) {
  if (-not (Test-IsAdmin)) {
    $actions += "ERROR:NotAdmin"
  } elseif ($denied -contains "SecureBootDisabled") {
    $actions += "NoChange:SecureBootDisabled"
  } elseif ($updated) {
    $actions += "NoChange:AlreadyUpdated"
  } else {
    if ($highConfidenceOptOut -eq 1) {
      # Set to 0 or remove; removing is cleanest (default is opt-in)
      Remove-RegValue $sbRoot "HighConfidenceOptOut"
      $actions += "Removed HighConfidenceOptOut"
    }

    if ($null -eq $availableUpdates -or $availableUpdates -eq 0) {
      Set-RegDword $sbRoot "AvailableUpdates" 0x5944
      $actions += "Set AvailableUpdates=0x5944"
    } else {
      $actions += ("NoChange:AvailableUpdatesAlreadySet(0x{0})" -f $availableUpdates.ToString("X"))
    }

    try {
      Start-ScheduledTask -TaskName $task
      $actions += "Started task: $task"
    } catch {
      $actions += "WARN:CouldNotStartTask"
    }
  }
}

[pscustomobject]@{
  ComputerName             = $env:COMPUTERNAME
  UEFISecureBootEnabledReg = $uefiSecureBootEnabled
  ConfirmSecureBootUEFI    = $confirmSecureBoot
  UEFICA2023Status         = $uefiStatus
  UEFICA2023Error          = $uefiError
  HighConfidenceOptOut     = $highConfidenceOptOut
  MicrosoftUpdateManagedOptIn = $msManagedOptIn
  AvailableUpdates         = if ($null -eq $availableUpdates) { $null } else { ("0x{0}" -f $availableUpdates.ToString("X")) }
  Event1808Latest = if ($ev1808) { $ev1808 | Select-Object -First 1 -ExpandProperty TimeCreated } else { $null }
  Event1801Latest = if ($ev1801) { $ev1801 | Select-Object -First 1 -ExpandProperty TimeCreated } else { $null }
  Updated                  = $updated
  DeniedReasons            = if ($denied.Count -eq 0) { $null } else { $denied -join ";" }
  Recommended              = if ($recommended.Count -eq 0) { "None" } else { $recommended -join "; " }
  Actions                  = if ($actions.Count -eq 0) { $null } else { $actions -join "; " }
}