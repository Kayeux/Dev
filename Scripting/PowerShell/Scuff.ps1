<#
.SYNOPSIS
  Detects GPOs that configure Device Guard / VBS / HVCI settings which can
  conflict with Intune Security Baselines.
 
.DESCRIPTION
  Queries RSoP (WMI) under root\rsop\computer:
   - RSOP_RegistryPolicySetting: which registry keys/values were set by GPO
   - RSOP_GPO: map GPOID/GUID to friendly GPO names
 
  Outputs a structured report (console + JSON) and returns exit code:
    0 = No conflicts found
    1 = Conflicts found
    2 = RSOP data unavailable or error (see message)
 
.NOTES
  Run as SYSTEM or Local Administrator.
  RSOP classes: https://learn.microsoft.com/.../rsop-registrypolicysetting
  GPO class:    https://wutils.com/wmi/root/rsop/computer/rsop_gpo/
#>
 
[CmdletBinding()]
param(
  # Where to save the JSON report
  [string]$OutJson = "C:\ProgramData\BaselineGpoConflicts.json",
 
  # Additional registry paths (prefix match) to inspect for GPO-applied settings
  [string[]]$ExtraRegPrefixes = @()
)
 
function Test-IsSystem {
  try {
    return ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq "S-1-5-18")
  } catch { return $false }
}
 
Write-Host "=== Baseline GPO Conflict Check (Computer scope) ==="
 
# --- Target registry prefixes & value names we care about ---
$targets = @(
  @{ Area="VBS";   RegPrefix='HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard'; ValueName='EnableVirtualizationBasedSecurity' },
  @{ Area="VBS";   RegPrefix='HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard'; ValueName='RequirePlatformSecurityFeatures' },
  @{ Area="HVCI";  RegPrefix='HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'; ValueName='Enabled' },
  @{ Area="HVCI";  RegPrefix='HKLM\SYSTEM\CurrentControlSet\Control\DeviceGuard\HypervisorEnforcedCodeIntegrity';            ValueName='Enabled' } # legacy path
)
 
# Merge any extra prefixes provided
foreach ($p in $ExtraRegPrefixes) {
  $targets += @{ Area="Custom"; RegPrefix=$p; ValueName=$null }
}
 
# --- Ensure RSOP namespace is available ---
$namespace = "root\rsop\computer"
try {
  $null = Get-CimInstance -Namespace $namespace -ClassName RSOP_Session -ErrorAction Stop
} catch {
  Write-Warning "RSOP data for COMPUTER scope isn't available yet. Try: gpupdate /force, then re-run."
  Write-Warning "If you run as standard user, COMPUTER-scope RSOP won't enumerate. Run as SYSTEM/admin."
  # Exit with 2 to indicate 'indeterminate'
  exit 2
}
 
# --- Pull all RegistryPolicySetting entries once ---
$regPolicies = Get-CimInstance -Namespace $namespace -ClassName RSOP_RegistryPolicySetting -ErrorAction SilentlyContinue
 
# --- Pull GPO table to map GPOID/GUID -> friendly name ---
$gpoMap = @{}
try {
  $gpos = Get-CimInstance -Namespace $namespace -ClassName RSOP_GPO -ErrorAction Stop
  foreach ($g in $gpos) {
    # RSOP_GPO has properties: name (friendly), guidName, id, fileSystemPath, etc.
    if ($g.guidName) { $gpoMap[$g.guidName.ToString().Trim('{}')] = $g.name }
  }
} catch {
  Write-Verbose "Failed to enumerate RSOP_GPO. Will show GUIDs only."
}
 
# --- Helper to decode REG value bytes based on valueType ---
function Convert-PolicyValue {
  param(
    [byte[]]$Bytes,
    [uint32]$Type
  )
  switch ($Type) {
    1 { # REG_SZ
      try { return [System.Text.Encoding]::Unicode.GetString($Bytes).TrimEnd([char]0) } catch { return "[REG_SZ]" }
    }
    2 { # EXPAND_SZ
      try { return [System.Text.Encoding]::Unicode.GetString($Bytes).TrimEnd([char]0) } catch { return "[REG_EXPAND_SZ]" }
    }
    3 { # REG_BINARY
      return ("0x" + ($Bytes | ForEach-Object { $_.ToString("X2") }) -join "")
    }
    4 { # REG_DWORD
      if ($Bytes.Length -ge 4) { return [BitConverter]::ToUInt32($Bytes,0) } else { return "[DWORD?]" }
    }
    7 { # REG_MULTI_SZ
      try {
        $str = [System.Text.Encoding]::Unicode.GetString($Bytes)
        return ($str -split "\x00" | Where-Object { $_ -ne "" })
      } catch { return "[REG_MULTI_SZ]" }
    }
    11 { # REG_QWORD
      if ($Bytes.Length -ge 8) { return [BitConverter]::ToUInt64($Bytes,0) } else { return "[QWORD?]" }
    }
    default { return "[Type:$Type Bytes:$($Bytes.Length)]" }
  }
}
 
# --- Match RSOP entries against our target list ---
$results = New-Object System.Collections.Generic.List[object]
 
foreach ($t in $targets) {
  $prefix = ($t.RegPrefix -replace '^HKLM\\', 'Machine\') # RSOP stores hives as "Machine\..." not HKLM\
  $valName = $t.ValueName
 
  $matches = $regPolicies | Where-Object {
    # Example RSOP registryKey: "Machine\System\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
    $_.registryKey -like "$prefix*" -and `
      ( [string]::IsNullOrEmpty($valName) -or $_.valueName -ieq $valName )
  }
 
  foreach ($m in $matches) {
    # RSOP_RegistryPolicySetting has: GPOID (LDAP path incl. GUID), value (bytes), valueType (DWORD, etc.), precedence
    $guid = ($m.GPOID -replace '.*\{','' -replace '\}.*','')
    $gpoFriendly = $null
    if ($guid -and $gpoMap.ContainsKey($guid)) { $gpoFriendly = $gpoMap[$guid] } else { $gpoFriendly = "{${guid}}" }
 
    $decoded = Convert-PolicyValue -Bytes $m.value -Type $m.valueType
 
    $results.Add([pscustomobject]@{
      Area        = $t.Area
      RegistryKey = $m.registryKey
      ValueName   = $m.valueName
      ValueType   = $m.valueType
      Value       = $decoded
      GPOGuid     = $guid
      GPOName     = $gpoFriendly
      Precedence  = $m.precedence
      # Any RSOP-registry policy set by a GPO is considered "conflicting" for Intune baseline purposes
      Conflict    = $true
    })
  }
}
 
# --- Condense and output ---
$conflicts = $results | Sort-Object Area, RegistryKey, ValueName, Precedence
 
if ($conflicts.Count -eq 0) {
  Write-Host "✅ No GPO-applied registry policies found for the monitored baseline keys (computer scope)."
} else {
  Write-Host "❗ Found $($conflicts.Count) GPO-applied setting(s) that can conflict with Intune baselines:`n"
  $conflicts | Select-Object Area, GPOName, Precedence, RegistryKey, ValueName, Value |
    Format-Table -AutoSize
}
 
# --- Save JSON report ---
try {
  $dir = Split-Path -Path $OutJson -Parent
  if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
  $meta = [pscustomobject]@{
    RanAsSystem = (Test-IsSystem)
    TimeUTC     = [DateTime]::UtcNow
    Hostname    = $env:COMPUTERNAME
    Namespace   = $namespace
    Conflicts   = $conflicts
  }
  $meta | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutJson -Encoding UTF8
  Write-Host "`nSaved JSON report to: $OutJson"
} catch {
  Write-Warning "Could not write JSON report: $($_.Exception.Message)"
}
 
# --- Exit codes for automation (e.g., Intune detection/remediation) ---
if ($conflicts.Count -gt 0) { exit 1 } else { exit 0 }