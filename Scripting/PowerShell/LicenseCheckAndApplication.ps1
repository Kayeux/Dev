<#
.SYNOPSIS
  Audit/remediate licensing: remove one SKU (e.g., AAD_PREMIUM_P1) and ensure another (e.g., Business Premium).

.PREREQS
  Microsoft Graph PowerShell SDK installed on the machine running this script.

.NOTES
  - Flags group-based assignments using licenseAssignmentStates.assignedByGroup.
  - Use -Mode Audit first, then -Mode Apply -WhatIf, then -Mode Apply to commit.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [object[]]$Tenants,

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$RemoveSkuPartNumber = "AAD_PREMIUM_P1",

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$EnsureSkuPartNumber = "SPB",

  [Parameter()]
  [ValidateSet("Audit","Apply")]
  [string]$Mode = "Audit",

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$OutDir = ".\LicenseAudit",

  [Parameter()]
  [switch]$IncludeDisabledUsers
)

begin {
  # TLS 1.2 for Windows PowerShell 5.1
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    throw "Microsoft.Graph PowerShell SDK not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
  }

  Import-Module Microsoft.Graph -ErrorAction Stop

  if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
  }

  $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
}

process {
  foreach ($t in $Tenants) {
    $tenantName = $t.Name
    $tenantId   = $t.TenantId

    if ([string]::IsNullOrWhiteSpace($tenantName) -or [string]::IsNullOrWhiteSpace($tenantId)) {
      Write-Warning "Skipping tenant entry missing Name or TenantId: $($t | ConvertTo-Json -Compress)"
      continue
    }

    Write-Host "=== Tenant: $tenantName ($tenantId) ==="

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    # Least-privileged permission for assignLicense is LicenseAssignment.ReadWrite.All
    # Add Organization.Read.All for subscribed SKUs, and User.Read.All to enumerate users.
    $scopes = @(
      "LicenseAssignment.ReadWrite.All",
      "Organization.Read.All",
      "User.Read.All"
    )

    Connect-MgGraph -TenantId $tenantId -Scopes $scopes -NoWelcome

    $skus = Get-MgSubscribedSku -All

    $removeSku = $skus | Where-Object { $_.SkuPartNumber -eq $RemoveSkuPartNumber } | Select-Object -First 1
    if (-not $removeSku) {
      Write-Warning "Remove SKU '$RemoveSkuPartNumber' not found in $tenantName. Exporting available SKUs and skipping."
      $skus | Select-Object SkuPartNumber,SkuId,ConsumedUnits,@{n="EnabledUnits";e={$_.PrepaidUnits.Enabled}} |
        Sort-Object SkuPartNumber |
        Export-Csv -NoTypeInformation -Path (Join-Path $OutDir "$tenantName`_SKUs_$timestamp.csv")
      continue
    }

    $ensureSku = $skus | Where-Object { $_.SkuPartNumber -eq $EnsureSkuPartNumber } | Select-Object -First 1
    if (-not $ensureSku) {
      Write-Warning "Ensure SKU '$EnsureSkuPartNumber' not found in $tenantName. Exporting available SKUs and skipping."
      $skus | Select-Object SkuPartNumber,SkuId,ConsumedUnits,@{n="EnabledUnits";e={$_.PrepaidUnits.Enabled}} |
        Sort-Object SkuPartNumber |
        Export-Csv -NoTypeInformation -Path (Join-Path $OutDir "$tenantName`_SKUs_$timestamp.csv")
      continue
    }

    $ensureAvailable = [int]($ensureSku.PrepaidUnits.Enabled) - [int]($ensureSku.ConsumedUnits)
    Write-Host ("Ensure SKU {0} available: {1}" -f $EnsureSkuPartNumber, $ensureAvailable)

    $props = "id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses,licenseAssignmentStates"
    $users = Get-MgUser -All -Property $props

    if (-not $IncludeDisabledUsers) {
      $users = $users | Where-Object { $_.AccountEnabled -eq $true }
    }
    $users = $users | Where-Object { $_.UserType -eq "Member" }

    $report = New-Object System.Collections.Generic.List[object]

    foreach ($u in $users) {
      $assignedSkuIds = @()
      if ($u.AssignedLicenses) { $assignedSkuIds = @($u.AssignedLicenses | ForEach-Object { $_.SkuId }) }

      $hasRemove = $assignedSkuIds -contains $removeSku.SkuId
      if (-not $hasRemove) { continue }

      $hasEnsure = $assignedSkuIds -contains $ensureSku.SkuId

      $las = @()
      if ($u.LicenseAssignmentStates) {
        $las = @($u.LicenseAssignmentStates | Where-Object { $_.SkuId -eq $removeSku.SkuId })
      }

      $assignedByGroupIds = @(
        $las | Where-Object { $_.AssignedByGroup } |
        ForEach-Object { $_.AssignedByGroup } | Select-Object -Unique
      )
      $isGroupBased = ($assignedByGroupIds.Count -gt 0)

      $action = if ($isGroupBased) {
        "SKIP_GROUP_BASED"
      } elseif ($hasEnsure) {
        "REMOVE_ONLY"
      } else {
        "ADD_ENSURE_AND_REMOVE"
      }

      $report.Add([pscustomobject]@{
        Tenant              = $tenantName
        UserPrincipalName   = $u.UserPrincipalName
        DisplayName         = $u.DisplayName
        RemoveSku           = $RemoveSkuPartNumber
        EnsureSku           = $EnsureSkuPartNumber
        HasEnsureAlready    = $hasEnsure
        RemoveIsGroupBased  = $isGroupBased
        AssignedByGroupIds  = ($assignedByGroupIds -join ";")
        PlannedAction       = $action
        Result              = ""
        Error               = ""
      }) | Out-Null
    }

    $tenantReportPath = Join-Path $OutDir "$tenantName`_Audit_$timestamp.csv"
    $report | Export-Csv -NoTypeInformation -Path $tenantReportPath
    Write-Host "Wrote audit: $tenantReportPath (rows: $($report.Count))"

    if ($Mode -ne "Apply") { continue }

    $needsEnsure = @($report | Where-Object { $_.PlannedAction -eq "ADD_ENSURE_AND_REMOVE" }).Count
    if ($needsEnsure -gt $ensureAvailable) {
      Write-Warning "Not enough '$EnsureSkuPartNumber' available ($ensureAvailable) for $needsEnsure users. Aborting apply for $tenantName."
      continue
    }

    foreach ($row in $report) {
      if ($row.PlannedAction -eq "SKIP_GROUP_BASED") {
        $row.Result = "Skipped (group-based). Fix group licensing instead."
        continue
      }

      $userId = $row.UserPrincipalName
      $addLicenses = @()
      $removeLicenses = @($removeSku.SkuId)

      if ($row.PlannedAction -eq "ADD_ENSURE_AND_REMOVE") {
        $addLicenses = @(@{ SkuId = $ensureSku.SkuId })
      }

      $target = "$tenantName\: $userId"
      if ($PSCmdlet.ShouldProcess($target, "Set-MgUserLicense add=$($addLicenses.Count) remove=1")) {
        try {
          Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses $removeLicenses | Out-Null
          $row.Result = "OK"
        } catch {
          $row.Result = "FAILED"
          $row.Error  = $_.Exception.Message
        }
      } else {
        $row.Result = "WhatIf/Confirm declined"
      }
    }

    $applyPath = Join-Path $OutDir "$tenantName`_Apply_$timestamp.csv"
    $report | Export-Csv -NoTypeInformation -Path $applyPath
    Write-Host "Wrote apply results: $applyPath"
  }
}

end {
  Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
