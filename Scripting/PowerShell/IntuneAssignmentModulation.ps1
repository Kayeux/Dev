<#
.SYNOPSIS
Bulk assign Intune apps across one or more tenants using Microsoft Graph (delegated auth).

SAFE DEFAULTS
- Targets ALL Intune apps if -AppListPath is omitted
- Intent defaults to AVAILABLE (NOT REQUIRED)
- Skips apps that already have any assignments (unless overridden)
- Supports -MaxAssignmentsPerTenant to cap POST attempts per tenant (blast radius control)
- Supports -TargetGroup (GUID or exact display name). Default is All Users (all licensed users)
- -WhatIf prevents assignment POSTs but still performs GETs (needed to simulate safely)

REQUIREMENTS
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

Delegated scopes typically required:
- DeviceManagementApps.ReadWrite.All
- DeviceManagementConfiguration.ReadWrite.All
Optional (only when -TargetGroup is used):
- Group.Read.All

LOGGING
Creates:
.\IntuneAppAssignLogs\Run_yyyyMMdd_HHmmss\
  Results.csv
  Results.json
  Summary.json
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$TenantIds,

    # OPTIONAL: if omitted, targets ALL Intune apps in each tenant
    [string]$AppListPath,

    # Per your request: default AVAILABLE (do not pass -Intent unless you truly mean it)
    [ValidateSet('available','required','uninstall','availableWithoutEnrollment')]
    [string]$Intent = 'available',

    # Only used when AppListPath is provided
    [ValidateSet('Exact','Contains')]
    [string]$MatchMode = 'Exact',

    # Only used when AppListPath is provided
    [switch]$AllowContainsFallback,

    [string]$LogRoot = ".\IntuneAppAssignLogs",

    [switch]$UseBetaForAssignmentFallback,

    # 0 = no limit (default). Otherwise limit assignment POST attempts per tenant.
    [ValidateRange(0,2147483647)]
    [int]$MaxAssignmentsPerTenant = 0,

    # If set, do not skip just because app already has assignments.
    [switch]$ReassignEvenIfAlreadyAssigned,

    # Default = All Users (all licensed users). If provided, assign to this group instead.
    # Accepts GUID or exact display name (resolved in each tenant).
    [string]$TargetGroup,

    # When targeting ALL apps, make ordering deterministic so "first 10" is predictable.
    [ValidateSet('DisplayName','None')]
    [string]$AllAppsOrderBy = 'DisplayName'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-GraphSdkCommands {
    $needed = @('Connect-MgGraph','Disconnect-MgGraph','Invoke-MgGraphRequest')
    foreach ($cmd in $needed) {
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
            throw "Missing required command '$cmd'. Install module: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
        }
    }
}

function Escape-ODataString {
    param([Parameter(Mandatory = $true)][string]$Value)
    return $Value.Replace("'", "''")
}

function Test-IsGuid {
    param([Parameter(Mandatory = $true)][string]$Value)
    $tmp = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$tmp)
}

function Read-AppNamesFromFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "App list file not found: $Path"
    }

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $names = New-Object System.Collections.Generic.List[string]

    switch ($ext) {
        '.txt' {
            Get-Content -LiteralPath $Path | ForEach-Object {
                $line = $_.Trim()
                if ($line -and -not $line.StartsWith('#')) {
                    [void]$names.Add($line)
                }
            }
        }

        '.csv' {
            $rows = Import-Csv -LiteralPath $Path
            if (-not $rows) { throw "CSV is empty: $Path" }

            foreach ($row in $rows) {
                $candidate = $null

                # Common columns you might see / want to use
                foreach ($col in @('AppName','DisplayName','Name','appName','displayName')) {
                    if ($row.PSObject.Properties.Name -contains $col) {
                        $candidate = [string]$row.$col
                        if (-not [string]::IsNullOrWhiteSpace($candidate)) { break }
                    }
                }

                # Fallback: first non-empty column
                if ([string]::IsNullOrWhiteSpace($candidate)) {
                    foreach ($p in $row.PSObject.Properties) {
                        if ($null -ne $p.Value -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
                            $candidate = [string]$p.Value
                            break
                        }
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                    [void]$names.Add($candidate.Trim())
                }
            }
        }

        default {
            throw "Unsupported file type '$ext'. Use .txt or .csv"
        }
    }

    $deduped = $names |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique

    if (-not $deduped -or $deduped.Count -eq 0) {
        throw "No app names found in $Path"
    }

    return ,$deduped
}

function New-RunFolder {
    param([Parameter(Mandatory = $true)][string]$Root)

    # IMPORTANT: Use .NET so folder creation is NOT suppressed by -WhatIf
    $null = [System.IO.Directory]::CreateDirectory($Root)

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $folder = Join-Path $Root "Run_$stamp"
    $null = [System.IO.Directory]::CreateDirectory($folder)

    return [System.IO.Path]::GetFullPath($folder)
}

function Get-ObjPropString {
    param(
        [Parameter(Mandatory=$true)]$Obj,
        [Parameter(Mandatory=$true)][string]$Name,
        [string]$Default = ''
    )

    if ($null -eq $Obj) { return $Default }

    $p = $Obj.PSObject.Properties[$Name]
    if ($null -ne $p -and $null -ne $p.Value) {
        return [string]$p.Value
    }

    return $Default
}

function Add-ResultRecord {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$List,

        [Parameter(Mandatory = $true)]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [string]$RequestedApp,

        [string]$MatchedApp,
        [string]$AppId,
        [string]$AppODataType,
        [string]$Result,
        [string]$Details
    )

    $List.Add([pscustomobject]@{
        Timestamp    = (Get-Date).ToString("s")
        TenantId     = $TenantId
        RequestedApp = $RequestedApp
        MatchedApp   = $MatchedApp
        AppId        = $AppId
        AppType      = $AppODataType
        Result       = $Result
        Details      = $Details
    }) | Out-Null
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET','POST')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [object]$Body = $null,

        [int]$MaxRetries = 4
    )

    $attempt = 0
    do {
        try {
            $attempt++

            if ($Method -eq 'GET') {
                return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
            }

            if ($null -eq $Body) {
                return Invoke-MgGraphRequest -Method POST -Uri $Uri -OutputType PSObject
            }

            $json = $Body | ConvertTo-Json -Depth 50 -Compress
            return Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $json -ContentType 'application/json' -OutputType PSObject
        }
        catch {
            $raw = ($_ | Out-String)
            $isRetryable = ($raw -match '\b429\b') -or ($raw -match '\b503\b') -or ($raw -match '\b504\b')

            if ($attempt -lt $MaxRetries -and $isRetryable) {
                $sleep = [Math]::Min(30, [int][Math]::Pow(2, $attempt))
                Start-Sleep -Seconds $sleep
                continue
            }

            throw
        }
    } while ($true)
}

function Get-GraphPaged {
    param([Parameter(Mandatory = $true)][string]$InitialUri)

    $all  = New-Object System.Collections.Generic.List[object]
    $next = $InitialUri

    while ($next) {
        $resp = Invoke-GraphRequestWithRetry -Method GET -Uri $next

        # StrictMode-safe: check property existence rather than touching missing properties
        $valueProp = $resp.PSObject.Properties['value']
        if ($valueProp -and $null -ne $valueProp.Value) {
            foreach ($item in @($valueProp.Value)) { $all.Add($item) | Out-Null }

            $nextProp = $resp.PSObject.Properties['@odata.nextLink']
            if ($nextProp -and $nextProp.Value) {
                $next = [string]$nextProp.Value
            } else {
                $next = $null
            }
        }
        else {
            # Single-object response (no paging)
            $all.Add($resp) | Out-Null
            $next = $null
        }
    }

    return ,$all.ToArray()
}

function Find-MatchingApps {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Apps,

        [Parameter(Mandatory = $true)]
        [string]$RequestedName,

        [ValidateSet('Exact','Contains')]
        [string]$MatchMode = 'Exact',

        [switch]$AllowContainsFallback
    )

    $requested = $RequestedName.Trim()
    $requestedLower = $requested.ToLowerInvariant()

    $exact = @(
        $Apps | Where-Object {
            $_.displayName -and $_.displayName.Trim().ToLowerInvariant() -eq $requestedLower
        }
    )
    if ($exact.Count -gt 0) { return ,$exact }

    if ($MatchMode -eq 'Contains' -or $AllowContainsFallback) {
        $contains = @(
            $Apps | Where-Object {
                $_.displayName -and $_.displayName.ToLowerInvariant().Contains($requestedLower)
            }
        )
        return ,$contains
    }

    return @()
}

function Resolve-AssignmentTarget {
    param([string]$TargetGroup)

    if ([string]::IsNullOrWhiteSpace($TargetGroup)) {
        return [pscustomobject]@{
            Mode         = 'AllUsers'
            DisplayLabel = 'All Users (all licensed users)'
            TargetObject = @{
                '@odata.type' = '#microsoft.graph.allLicensedUsersAssignmentTarget'
            }
        }
    }

    $groupId = $null
    $groupDisplayName = $null

    if (Test-IsGuid -Value $TargetGroup) {
        $g = Invoke-GraphRequestWithRetry -Method GET -Uri "/v1.0/groups/$TargetGroup?`$select=id,displayName"
        $groupId = [string]$g.id
        $groupDisplayName = [string]$g.displayName
    }
    else {
        $escaped = Escape-ODataString -Value $TargetGroup
        $resp = Invoke-GraphRequestWithRetry -Method GET -Uri "/v1.0/groups?`$filter=displayName eq '$escaped'&`$select=id,displayName&`$top=10"
        $matches = @($resp.value)

        if ($matches.Count -eq 0) {
            throw "TargetGroup display name '$TargetGroup' was not found in this tenant."
        }
        if ($matches.Count -gt 1) {
            $names = ($matches | ForEach-Object { '{0} ({1})' -f $_.displayName, $_.id }) -join '; '
            throw "TargetGroup display name '$TargetGroup' is ambiguous in this tenant. Matches: $names. Use the group GUID instead."
        }

        $groupId = [string]$matches[0].id
        $groupDisplayName = [string]$matches[0].displayName
    }

    return [pscustomobject]@{
        Mode         = 'Group'
        DisplayLabel = ('Group: {0} ({1})' -f $groupDisplayName, $groupId)
        TargetObject = @{
            '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
            groupId       = $groupId
        }
        GroupId       = $groupId
        GroupName     = $groupDisplayName
    }
}

function Get-AssignmentTargetFingerprint {
    param([Parameter(Mandatory = $true)][object]$AssignmentTarget)

    $odata = [string]$AssignmentTarget.'@odata.type'
    switch ($odata) {
        '#microsoft.graph.allLicensedUsersAssignmentTarget' { return 'allLicensedUsers' }
        'microsoft.graph.allLicensedUsersAssignmentTarget'  { return 'allLicensedUsers' }
        '#microsoft.graph.groupAssignmentTarget'            { return ('group:{0}' -f [string]$AssignmentTarget.groupId) }
        'microsoft.graph.groupAssignmentTarget'             { return ('group:{0}' -f [string]$AssignmentTarget.groupId) }
        default {
            if ($AssignmentTarget.PSObject.Properties.Name -contains 'groupId') {
                return ('other:{0}:{1}' -f $odata, [string]$AssignmentTarget.groupId)
            }
            return ('other:{0}' -f $odata)
        }
    }
}

function Test-AssignmentAlreadyHasSameTargetAndIntent {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Assignments,

        [Parameter(Mandatory = $true)]
        [string]$Intent,

        [Parameter(Mandatory = $true)]
        [hashtable]$DesiredTarget
    )

    $desiredFingerprint = Get-AssignmentTargetFingerprint -AssignmentTarget $DesiredTarget

    foreach ($a in $Assignments) {
        $aIntent = [string]$a.intent
        $aTarget = $a.target
        if ($null -eq $aTarget) { continue }

        $fp = Get-AssignmentTargetFingerprint -AssignmentTarget $aTarget
        if ($aIntent -eq $Intent -and $fp -eq $desiredFingerprint) {
            return $true
        }
    }

    return $false
}

function Get-MobileAppAssignmentSettingsBody {
    param([Parameter(Mandatory = $true)][string]$AppODataType)

    $fallback = @{
        '@odata.type' = '#microsoft.graph.mobileAppAssignmentSettings'
    }

    switch -Regex ($AppODataType) {
        '^#microsoft\.graph\.win32LobApp$' {
            return @{
                '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings'
                notifications = 'showAll'
            }
        }
        '^#microsoft\.graph\.winGetApp$' {
            return @{ '@odata.type' = '#microsoft.graph.winGetAppAssignmentSettings' }
        }
        '^#microsoft\.graph\.windowsUniversalAppX$' {
            return @{
                '@odata.type' = '#microsoft.graph.windowsUniversalAppXAppAssignmentSettings'
                useDeviceContext = $false
            }
        }
        '^#microsoft\.graph\.microsoftStoreForBusinessApp$' {
            return @{
                '@odata.type' = '#microsoft.graph.windowsUniversalAppXAppAssignmentSettings'
                useDeviceContext = $false
            }
        }
        '^#microsoft\.graph\.iosVppApp$' {
            return @{
                '@odata.type' = '#microsoft.graph.iosVppAppAssignmentSettings'
                useDeviceLicensing = $false
            }
        }
        default { return $fallback }
    }
}

function New-MobileAppAssignmentBody {
    param(
        [Parameter(Mandatory = $true)][string]$Intent,
        [Parameter(Mandatory = $true)][string]$AppODataType,
        [Parameter(Mandatory = $true)][hashtable]$TargetObject
    )

    @{
        '@odata.type' = '#microsoft.graph.mobileAppAssignment'
        intent   = $Intent
        target   = $TargetObject
        settings = (Get-MobileAppAssignmentSettingsBody -AppODataType $AppODataType)
    }
}

function Try-CreateAppAssignment {
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$Intent,
        [Parameter(Mandatory = $true)][string]$AppODataType,
        [Parameter(Mandatory = $true)][hashtable]$TargetObject,
        [switch]$UseBetaForAssignmentFallback
    )

    $body = New-MobileAppAssignmentBody -Intent $Intent -AppODataType $AppODataType -TargetObject $TargetObject

    try {
        [void](Invoke-GraphRequestWithRetry -Method POST -Uri "/v1.0/deviceAppManagement/mobileApps/$AppId/assignments" -Body $body)
        return [pscustomobject]@{ Success = $true; Message = "Created assignment via v1.0 /assignments" }
    }
    catch {
        $firstErr = ($_ | Out-String)

        if ($UseBetaForAssignmentFallback) {
            try {
                [void](Invoke-GraphRequestWithRetry -Method POST -Uri "/beta/deviceAppManagement/mobileApps/$AppId/assignments" -Body $body)
                return [pscustomobject]@{ Success = $true; Message = "Created assignment via beta /assignments fallback" }
            }
            catch {
                $secondErr = ($_ | Out-String)
                return [pscustomobject]@{
                    Success = $false
                    Message = ("v1.0 failed: {0}`nBeta fallback failed: {1}" -f $firstErr.Trim(), $secondErr.Trim())
                }
            }
        }

        return [pscustomobject]@{ Success = $false; Message = $firstErr.Trim() }
    }
}

# -------------------- Main --------------------

Test-GraphSdkCommands

$runFolder = New-RunFolder -Root $LogRoot
$results = New-Object 'System.Collections.Generic.List[object]'

$summary = [ordered]@{
    RunStarted                    = (Get-Date).ToString("s")
    Intent                        = $Intent
    TenantCount                   = $TenantIds.Count
    Mode                          = $(if ([string]::IsNullOrWhiteSpace($AppListPath)) { 'AllAppsInTenant' } else { 'AppListFile' })
    AppListPath                   = $AppListPath
    MaxAssignmentsPerTenant       = $MaxAssignmentsPerTenant
    ReassignEvenIfAlreadyAssigned = [bool]$ReassignEvenIfAlreadyAssigned
    AssignmentTargetRequested     = $(if ([string]::IsNullOrWhiteSpace($TargetGroup)) { 'All Users (all licensed users)' } else { $TargetGroup })
    Assigned                      = 0
    AlreadyAssignedSkipped        = 0
    AlreadyHasSameTargetSkipped   = 0
    NotFound                      = 0
    Ambiguous                     = 0
    SkippedByLimit                = 0
    Errors                        = 0
    WhatIf                        = [bool]$WhatIfPreference
    RunFolder                     = $runFolder
}

Write-Host ""
Write-Host "Run folder: $($summary.RunFolder)"
Write-Host "Tenants: $($summary.TenantCount)"
Write-Host "Intent: $Intent"
Write-Host ("Target: {0}" -f $summary.AssignmentTargetRequested)
Write-Host ("Mode: {0}" -f $summary.Mode)
Write-Host ("ReassignEvenIfAlreadyAssigned: {0}" -f $summary.ReassignEvenIfAlreadyAssigned)
Write-Host ("MaxAssignmentsPerTenant: {0}" -f $summary.MaxAssignmentsPerTenant)
Write-Host ""

foreach ($tenantId in $TenantIds) {
    Write-Host ("========== Tenant: {0} ==========" -f $tenantId) -ForegroundColor Cyan

    $tenantAssignmentAttempts = 0
    $tenantLimitReached = $false

    try {
        try { Disconnect-MgGraph | Out-Null } catch {}

        $scopes = @(
            'DeviceManagementApps.ReadWrite.All',
            'DeviceManagementConfiguration.ReadWrite.All'
        )
        if (-not [string]::IsNullOrWhiteSpace($TargetGroup)) {
            $scopes += 'Group.Read.All'
        }

        Connect-MgGraph -TenantId $tenantId -Scopes $scopes -NoWelcome | Out-Null

        $resolvedTarget = Resolve-AssignmentTarget -TargetGroup $TargetGroup
        Write-Host ("Assignment target resolved: {0}" -f $resolvedTarget.DisplayLabel) -ForegroundColor Yellow

        $mobileApps = Get-GraphPaged -InitialUri "/v1.0/deviceAppManagement/mobileApps?`$top=200"
        if (-not $mobileApps -or $mobileApps.Count -eq 0) {
            throw "No mobile apps returned from Graph for tenant $tenantId"
        }

        if ($AllAppsOrderBy -eq 'DisplayName') {
            $mobileApps = $mobileApps | Sort-Object -Property displayName
        }

$workItems = New-Object System.Collections.Generic.List[object]

if (-not [string]::IsNullOrWhiteSpace($AppListPath)) {
    $requestedNames = Read-AppNamesFromFile -Path $AppListPath

    foreach ($requestedName in $requestedNames) {
        $matches = Find-MatchingApps -Apps $mobileApps -RequestedName $requestedName -MatchMode $MatchMode -AllowContainsFallback:$AllowContainsFallback

        if (-not $matches -or $matches.Count -eq 0) {
            Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $requestedName -Result 'NotFound' -Details 'No matching mobile app found'
            $summary.NotFound++
            continue
        }

        if ($matches.Count -gt 1) {
            $matchNames = ($matches | Select-Object -ExpandProperty displayName | Sort-Object -Unique) -join '; '
            Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $requestedName -Result 'Ambiguous' -Details ("Multiple matches: {0}" -f $matchNames)
            $summary.Ambiguous++
            continue
        }

        [void]$workItems.Add([pscustomobject]@{
            RequestedName = $requestedName
            App           = $matches[0]
        })
    }

    # IMPORTANT: if a list was provided but resolved to zero usable apps, stop this tenant
    if ($workItems.Count -eq 0) {
        throw "AppListPath was provided but none of the requested apps resolved uniquely. Nothing to do."
    }
}
    else {
    foreach ($app in $mobileApps) {
        if (-not $app.id -or -not $app.displayName) { continue }
        [void]$workItems.Add([pscustomobject]@{
            RequestedName = [string]$app.displayName
            App           = $app
        })
    }
}

        foreach ($item in $workItems) {

            if ($MaxAssignmentsPerTenant -gt 0 -and $tenantAssignmentAttempts -ge $MaxAssignmentsPerTenant) {
                $tenantLimitReached = $true
                Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp ([string]$item.App.displayName) -AppId ([string]$item.App.id) -AppODataType ([string]$item.App.'@odata.type') -Result 'SkippedByLimit' -Details ("MaxAssignmentsPerTenant ({0}) reached for tenant" -f $MaxAssignmentsPerTenant)
                $summary.SkippedByLimit++
                continue
            }

            try {
                $app = $item.App
                $appId = [string]$app.id
                $appName = [string]$app.displayName
                $appType = Get-ObjPropString -Obj $app -Name '@odata.type' -Default '#microsoft.graph.mobileApp'

                $assignmentResp = Invoke-GraphRequestWithRetry -Method GET -Uri "/v1.0/deviceAppManagement/mobileApps/$appId/assignments"
                $assignments = @($assignmentResp.value)
                $assignmentCount = $assignments.Count

                if (-not $ReassignEvenIfAlreadyAssigned -and $assignmentCount -gt 0) {
                    Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp $appName -AppId $appId -AppODataType $appType -Result 'AlreadyAssignedSkipped' -Details ("Assignments present: {0}" -f $assignmentCount)
                    $summary.AlreadyAssignedSkipped++
                    continue
                }

                if ($assignmentCount -gt 0) {
                    $alreadySameTargetIntent = Test-AssignmentAlreadyHasSameTargetAndIntent -Assignments $assignments -Intent $Intent -DesiredTarget $resolvedTarget.TargetObject
                    if ($alreadySameTargetIntent) {
                        Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp $appName -AppId $appId -AppODataType $appType -Result 'AlreadyHasSameTargetSkipped' -Details ("Assignment with same target + intent already exists. Existing assignments: {0}" -f $assignmentCount)
                        $summary.AlreadyHasSameTargetSkipped++
                        continue
                    }
                }

                if ($PSCmdlet.ShouldProcess("$tenantId :: $appName", ("Assign to {0} ({1})" -f $resolvedTarget.DisplayLabel, $Intent))) {

                    # Count POST attempts (blast radius control)
                    $tenantAssignmentAttempts++

                    $create = Try-CreateAppAssignment -AppId $appId -Intent $Intent -AppODataType $appType -TargetObject $resolvedTarget.TargetObject -UseBetaForAssignmentFallback:$UseBetaForAssignmentFallback

                    if ($create.Success) {
                        Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp $appName -AppId $appId -AppODataType $appType -Result 'Assigned' -Details ($create.Message + " | Target=" + $resolvedTarget.DisplayLabel)
                        $summary.Assigned++
                    }
                    else {
                        Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp $appName -AppId $appId -AppODataType $appType -Result 'Error' -Details ($create.Message + " | Target=" + $resolvedTarget.DisplayLabel)
                        $summary.Errors++
                    }
                }
                else {
                    Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp $appName -AppId $appId -AppODataType $appType -Result 'WhatIf' -Details ("Would assign to " + $resolvedTarget.DisplayLabel)
                }
            }
            catch {
                Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp $item.RequestedName -MatchedApp ([string]$item.App.displayName) -AppId ([string]$item.App.id) -AppODataType ([string]$item.App.'@odata.type') -Result 'Error' -Details (($_ | Out-String).Trim())
                $summary.Errors++
            }
        }

        if ($tenantLimitReached) {
            Write-Host ("Tenant assignment attempt cap reached ({0}). Remaining apps were logged as SkippedByLimit." -f $MaxAssignmentsPerTenant) -ForegroundColor DarkYellow
        }
    }
    catch {
        $msg = ($_ | Out-String).Trim()
        Write-Warning ("Tenant-level failure for {0}: {1}" -f $tenantId, $msg)
        Add-ResultRecord -List $results -TenantId $tenantId -RequestedApp '(TENANT)' -Result 'Error' -Details ("Tenant-level failure: {0}" -f $msg)
        $summary.Errors++
    }
    finally {
        try { Disconnect-MgGraph | Out-Null } catch {}
    }
}

$summary.RunEnded = (Get-Date).ToString("s")
$summary.TotalRecords = $results.Count

$resultsCsv = Join-Path $runFolder 'Results.csv'
$resultsJson = Join-Path $runFolder 'Results.json'
$summaryJson = Join-Path $runFolder 'Summary.json'

$results | Export-Csv -NoTypeInformation -Encoding UTF8 -LiteralPath $resultsCsv
$results | ConvertTo-Json -Depth 20 | Set-Content -Encoding UTF8 -LiteralPath $resultsJson
$summary | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -LiteralPath $summaryJson

Write-Host ""
Write-Host "========== SUMMARY ==========" -ForegroundColor Green
Write-Host ("Assigned:                    {0}" -f $summary.Assigned)
Write-Host ("AlreadyAssignedSkipped:      {0}" -f $summary.AlreadyAssignedSkipped)
Write-Host ("AlreadyHasSameTargetSkipped: {0}" -f $summary.AlreadyHasSameTargetSkipped)
Write-Host ("NotFound:                    {0}" -f $summary.NotFound)
Write-Host ("Ambiguous:                   {0}" -f $summary.Ambiguous)
Write-Host ("SkippedByLimit:              {0}" -f $summary.SkippedByLimit)
Write-Host ("Errors:                      {0}" -f $summary.Errors)
Write-Host ("Logs:                        {0}" -f $summary.RunFolder)
Write-Host "============================="
Write-Host ""

$hasIssues = ($summary.Errors -gt 0) -or
             ($summary.AlreadyAssignedSkipped -gt 0) -or
             ($summary.AlreadyHasSameTargetSkipped -gt 0) -or
             ($summary.Ambiguous -gt 0) -or
             ($summary.NotFound -gt 0)

if ($hasIssues) {
    $null = Read-Host ("Completed with issues. Errors={0}, AlreadyAssignedSkipped={1}, AlreadyHasSameTargetSkipped={2}, Ambiguous={3}, NotFound={4}, SkippedByLimit={5}. Review logs in: {6}. Press Enter to exit" -f `
        $summary.Errors, $summary.AlreadyAssignedSkipped, $summary.AlreadyHasSameTargetSkipped, $summary.Ambiguous, $summary.NotFound, $summary.SkippedByLimit, $summary.RunFolder)
}
else {
    $null = Read-Host ("Completed successfully. No errors/failures/already-assigned apps were found. SkippedByLimit={0}. Logs in: {1}. Press Enter to exit" -f `
        $summary.SkippedByLimit, $summary.RunFolder)
}