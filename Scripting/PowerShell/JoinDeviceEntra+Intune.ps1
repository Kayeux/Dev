<#
.SYNOPSIS
    Entra ID (Azure AD) Join & Intune Enrollment without scheduled tasks.
    Includes an AttemptFix parameter to optionally fix known issues.
    Restores original dsregcmd MDM detection (looking only for "MDM : Microsoft Intune"),
    AND adds a registry-based check for Intune.

.DESCRIPTION
    1) Optionally removes old auto-enrollment tasks from Task Scheduler, if any.
    2) Checks whether the device is already Azure AD joined. If not, tries dsregcmd /join.
       - If that fails, optionally tries debug/leave + registry cleanup, then a second /join attempt.
       - If still not joined, sets NinjaOne to "EntraFailed" and exits.
    3) Checks/creates needed MDM registry keys if -AttemptFix was specified (unchanged from prior version).
    4) Checks whether the device is already enrolled in Intune via:
       - dsregcmd output line: "MDM : Microsoft Intune" (the original approach),
       - or new registry-based logic under HKLM:\SOFTWARE\Microsoft\Enrollments.
       If either says “enrolled,” skip deviceenroller. Otherwise we attempt the appropriate deviceenroller command.
    5) Sets the NinjaOne custom field "entraIntuneJoinState" to "Entra," "Intune," "Both," or "None" (or "EntraFailed" if join fails).

NOTES:
    - We do NOT set MDMEnrolled = $true based on MdmUrl, since you said those URLs can be forcibly set, causing false positives.
    - If the line "MDM : Microsoft Intune" is missing, the script can still detect real Intune enrollment via the registry-based method.
#>

[CmdletBinding()]
param(
    [switch] $UseUserCredential,
    [switch] $DryRun,
    [switch] $AttemptFix
)

###############################################################################
# HELPER: Parse dsregcmd /status for relevant fields (restored old logic)
###############################################################################
function Get-DsregStatusFields {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DsregOutput
    )

    $results = @{
        "AzureAdJoined"       = $false
        "MDMEnrolled"         = $false
        "MdmUrl"              = $null
        "MdmTouUrl"           = $null
        "MdmComplianceUrl"    = $null
        "AzureAdPrtAuthority" = $null
        "TenantId"            = $null
    }

    # Check for Azure AD joined
    if ($DsregOutput -match "AzureAdJoined\s*:\s*YES") {
        $results["AzureAdJoined"] = $true
    }

    # ### RESTORED: The original check for "MDM : Microsoft Intune"
    # Some modern builds won't have this line, but we keep it for older or certain scenarios.
    if ($DsregOutput -match "MDM\s*:\s*Microsoft Intune") {
        $results["MDMEnrolled"] = $true
    }

    # We DO NOT set MDMEnrolled from MdmUrl now, just store it for logs:
    if ($DsregOutput -match "MdmUrl\s*:\s*(.+)") {
        $results["MdmUrl"] = $matches[1].Trim()
    }
    if ($DsregOutput -match "MdmTouUrl\s*:\s*(.+)") {
        $results["MdmTouUrl"] = $matches[1].Trim()
    }
    if ($DsregOutput -match "MdmComplianceUrl\s*:\s*(.+)") {
        $results["MdmComplianceUrl"] = $matches[1].Trim()
    }

    if ($DsregOutput -match "AzureAdPrtAuthority\s*:\s*(.+)") {
        $results["AzureAdPrtAuthority"] = $matches[1].Trim()
    }

    # Attempt to capture the TenantId field
    if ($DsregOutput -match "TenantId\s*:\s*([0-9a-fA-F-]+)") {
        $results["TenantId"] = $matches[1].Trim()
    }

    return $results
}

###############################################################################
# HELPER: Remove old auto-enrollment tasks (if they exist)
###############################################################################
function Remove-EnterpriseMgmtScheduledTasks {
    Write-Host "Searching for existing EnterpriseMgmt tasks..."

    $taskPath = "\Microsoft\Windows\EnterpriseMgmt"

    try {
        $tasks = Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue
        if ($null -ne $tasks -and $tasks.Count -gt 0) {
            Write-Host "Found $( $tasks.Count ) task(s) under $taskPath. Removing them..."
            foreach ($t in $tasks) {
                try {
                    Write-Host "Removing scheduled task '$($t.TaskName)' under '$taskPath'..."
                    Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $taskPath -Confirm:$false
                }
                catch {
                    Write-Warning "Could not remove scheduled task '$($t.TaskName)': $($_.Exception.Message)"
                }
            }
        }
        else {
            Write-Host "No scheduled tasks found under $taskPath."
        }
    }
    catch {
        Write-Warning "Could not enumerate tasks under $taskPath : $($_.Exception.Message)"
    }

    try {
        $subfolders = Get-ScheduledTask -TaskPath "$taskPath\" -ErrorAction SilentlyContinue |
                     Select-Object -ExpandProperty TaskPath -Unique |
                     Where-Object { $_ -ne $taskPath }

        foreach ($sf in $subfolders) {
            try {
                $subtasks = Get-ScheduledTask -TaskPath $sf -ErrorAction SilentlyContinue
                if ($null -ne $subtasks -and $subtasks.Count -gt 0) {
                    Write-Host "Removing tasks in subfolder '$sf'..."
                    foreach ($st in $subtasks) {
                        Unregister-ScheduledTask -TaskName $st.TaskName -TaskPath $sf -Confirm:$false
                    }
                }
                else {
                    Write-Host "No tasks found under subfolder '$sf'."
                }
            }
            catch {
                Write-Warning "Could not remove tasks from subfolder '$sf': $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Warning "Could not handle subfolders of $taskPath : $($_.Exception.Message)"
    }
}

###############################################################################
# HELPER: Ensure MDM registry keys exist under CloudDomainJoin\TenantInfo\<TenantId>
###############################################################################
function Ensure-MdmRegistryKeys {
    param(
        [string]$TenantId
    )

    if (-not $TenantId) {
        Write-Warning "No TenantId found, cannot check CloudDomainJoin registry keys."
        return
    }

    $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\$TenantId"

    Write-Host "Checking MDM registry keys under $basePath..."

    $defaults = @{
        "MdmEnrollmentUrl"  = "https://enrollment.manage.microsoft.com/enrollmentserver/discovery.svc"
        "MdmTermsOfUseUrl"  = "https://portal.manage.microsoft.com/TermsofUse.aspx"
        "MdmComplianceUrl"  = "https://portal.manage.microsoft.com/?portalAction=Compliance"
    }

    if (-not (Test-Path $basePath)) {
        Write-Warning "Tenant registry path $basePath does not exist!"
        return
    }

    foreach ($key in $defaults.Keys) {
        $regVal = (Get-ItemProperty -Path $basePath -Name $key -ErrorAction SilentlyContinue).$key
        if (-not $regVal) {
            Write-Warning "$key is missing or blank under $basePath."
        }
    }
}

function Fix-MdmRegistryKeys {
    param(
        [string]$TenantId
    )

    if (-not $TenantId) {
        Write-Warning "No TenantId found, cannot fix CloudDomainJoin registry keys."
        return
    }

    $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\$TenantId"
    $defaults = @{
        "MdmEnrollmentUrl"  = "https://enrollment.manage.microsoft.com/enrollmentserver/discovery.svc"
        "MdmTermsOfUseUrl"  = "https://portal.manage.microsoft.com/TermsofUse.aspx"
        "MdmComplianceUrl"  = "https://portal.manage.microsoft.com/?portalAction=Compliance"
    }

    if (Test-Path $basePath) {
        foreach ($key in $defaults.Keys) {
            $regVal = (Get-ItemProperty -Path $basePath -Name $key -ErrorAction SilentlyContinue).$key
            if (-not $regVal) {
                Write-Host "Setting missing $key to default '$($defaults[$key])'..."
                try {
                    Set-ItemProperty -Path $basePath -Name $key -Value $defaults[$key] -Type String
                }
                catch {
                    Write-Warning "Could not set $key : $($_.Exception.Message)"
                }
            }
            else {
                Write-Host "$key is already set to: $regVal"
            }
        }
    }
    else {
        Write-Warning "Tenant registry path $basePath does not exist; cannot fix MDM registry keys."
    }
}

###############################################################################
# HELPER: Set the NinjaOne field and exit the script
###############################################################################
function Fail-And-Exit {
    param(
        [string]$Reason = "EntraFailed"
    )

    Write-Warning "Azure AD join failure. Setting Ninja custom field to '$Reason' and exiting script."
    try {
        Ninja-Property-Set entraIntuneJoinState $Reason
        Write-Host "Updated NinjaOne custom field 'entraIntuneJoinState' to '$Reason'."
    }
    catch {
        Write-Warning "Could not set 'entraIntuneJoinState': $($_.Exception.Message)"
    }

    exit 1
}

###############################################################################
# ### NEW ### HELPER: Additional Registry-Based Check for Intune Enrollment
###############################################################################
function Test-LocalIntuneEnrollmentRegistry {
    <#
    Scans HKLM:\SOFTWARE\Microsoft\Enrollments for a subkey with:
        EnrollmentState = 1
        ProviderID matching "MS DM Server" or "Microsoft Device Management"
    If found, returns $true; otherwise $false.
    #>

    $enrollPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
    if (-not (Test-Path $enrollPath)) {
        return $false
    }

    Get-ChildItem -Path $enrollPath -ErrorAction SilentlyContinue |
        Where-Object { $_.PSIsContainer } |
        ForEach-Object {
            try {
                $props = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop
                $state = $props.EnrollmentState
                $provider = $props.ProviderID

                # If EnrollmentState=1 and ProviderID is typical for Intune
                if ($state -eq 1 -and ($provider -match "(?i)(MS DM Server|Microsoft Device Management)")) {
                    return $true
                }
            }
            catch {
                # ignore read errors
            }
        }

    return $false
}

###############################################################################
# 1. Remove old auto-enrollment tasks (if they exist)
###############################################################################
Write-Host "=== 1. Remove old auto-enrollment tasks (if they exist) ==="
if ($DryRun) {
    Write-Host "-DryRun specified; skipping scheduled task removal."
} else {
    Remove-EnterpriseMgmtScheduledTasks
}

###############################################################################
# 2. Check / Attempt Entra (Azure AD) Join
###############################################################################
Write-Host "=== 2. Check / Attempt Entra (Azure AD) Join ==="
Write-Host "Retrieving dsregcmd /status output..."
$initialDsregOutput = dsregcmd /status | Out-String
$fields = Get-DsregStatusFields -DsregOutput $initialDsregOutput

if ($fields.AzureAdJoined) {
    Write-Host "Device is already joined to Entra ID (Azure AD)."
}
else {
    Write-Host "Device is NOT joined to Entra ID."

    if ($DryRun) {
        Write-Host "-DryRun specified; skipping dsregcmd /join step."
    }
    else {
        Write-Host "Attempting dsregcmd /join..."
        try {
            & "$env:SystemRoot\System32\dsregcmd.exe" /join | Out-Null
        }
        catch {
            Write-Warning "dsregcmd /join encountered an error: $($_.Exception.Message)"
        }

        Write-Host "Waiting 15 seconds to allow the join to complete..."
        Start-Sleep -Seconds 15
    }

    $postJoinOutput = dsregcmd /status | Out-String
    $fields = Get-DsregStatusFields -DsregOutput $postJoinOutput

    if (-not $fields.AzureAdJoined) {
        Write-Warning "Initial dsregcmd /join did NOT succeed."
        if ($AttemptFix -and -not $DryRun) {
            Write-Host "-AttemptFix specified. Trying dsregcmd /debug /leave + registry cleanup."

            try {
                Write-Host "Running dsregcmd /debug /leave..."
                & "$env:SystemRoot\System32\dsregcmd.exe" /debug /leave | Out-Null
            }
            catch {
                Write-Warning "dsregcmd /debug /leave encountered an error: $($_.Exception.Message)"
            }

            Write-Host "Removing subkeys in HKLM:\SOFTWARE\Microsoft\Enrollments..."
            $enrollPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
            if (Test-Path $enrollPath) {
                Get-ChildItem -Path $enrollPath -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Write-Host "Removing key: $($_.PSPath)"
                        Remove-Item $_.PSPath -Recurse -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Failed to remove $($_.PSPath): $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Host "Enrollments registry path not found at $enrollPath."
            }

            Write-Host "Waiting 5 seconds after registry cleanup..."
            Start-Sleep -Seconds 5

            Write-Host "Attempting dsregcmd /join again..."
            try {
                & "$env:SystemRoot\System32\dsregcmd.exe" /join | Out-Null
            }
            catch {
                Write-Warning "dsregcmd /join encountered an error (second attempt): $($_.Exception.Message)"
            }

            Write-Host "Waiting 15 seconds after second join attempt..."
            Start-Sleep -Seconds 15

            $finalJoinOutput = dsregcmd /status | Out-String
            $fields = Get-DsregStatusFields -DsregOutput $finalJoinOutput

            if (-not $fields.AzureAdJoined) {
                Write-Warning "Device still not joined to Entra ID after AttemptFix steps."
                Fail-And-Exit -Reason "EntraFailed"
            }
            else {
                Write-Host "Device is now joined to Entra ID (Azure AD) after AttemptFix steps."
            }
        }
        else {
            Fail-And-Exit -Reason "EntraFailed"
        }
    }
}

###############################################################################
# 3. Check / Fix MDM registry keys (CloudDomainJoin) if needed
###############################################################################
Write-Host "=== 3. Check / Fix MDM registry keys (CloudDomainJoin) if needed ==="
Ensure-MdmRegistryKeys -TenantId $fields.TenantId
if ($AttemptFix -and -not $DryRun) {
    Fix-MdmRegistryKeys -TenantId $fields.TenantId
}

###############################################################################
# 4. Check / Attempt Intune Enrollment
###############################################################################
Write-Host "=== 4. Check / Attempt Intune Enrollment ==="

# ### NEW ### Combine old dsregcmd logic (looking for "MDM : Microsoft Intune") with registry-based logic
$alreadyIntuneByDsreg = $fields.MDMEnrolled
$alreadyIntuneByReg   = Test-LocalIntuneEnrollmentRegistry

if ($alreadyIntuneByDsreg -or $alreadyIntuneByReg) {
    Write-Host "Device appears to be enrolled in Intune (via dsregcmd or registry). Skipping deviceenroller."
}
else {
    Write-Host "Device is NOT enrolled in Intune by either method."

    # By default, device credentials; if -UseUserCredential, user credentials
    if ($UseUserCredential) {
        Write-Host "Using user credentials (deviceenroller.exe /c /AutoEnrollMDM)..."
        $enrollArgs = "/c /AutoEnrollMDM"
    }
    else {
        Write-Host "Using device credentials (deviceenroller.exe /c /AutoEnrollMDMUsingAADDeviceCredential)..."
        $enrollArgs = "/c /AutoEnrollMDMUsingAADDeviceCredential"
    }

    if ($DryRun) {
        Write-Host "-DryRun specified; skipping Intune enrollment command."
    }
    else {
        try {
            & "$env:SystemRoot\System32\deviceenroller.exe" $enrollArgs | Out-Null
        }
        catch {
            Write-Warning "deviceenroller encountered an error: $($_.Exception.Message)"
        }

        Write-Host "Waiting 15 seconds for Intune enrollment to (hopefully) complete..."
        Start-Sleep -Seconds 15
    }

    # Re-check dsregcmd & registry
    $postEnrollOutput = dsregcmd /status | Out-String
    $fields = Get-DsregStatusFields -DsregOutput $postEnrollOutput
    $alreadyIntuneByDsreg = $fields.MDMEnrolled
    $alreadyIntuneByReg   = Test-LocalIntuneEnrollmentRegistry

    if ($alreadyIntuneByDsreg -or $alreadyIntuneByReg) {
        Write-Host "Device successfully enrolled in Intune (confirmed via dsregcmd or registry)."
    }
    else {
        Write-Warning "Intune enrollment not confirmed by either dsregcmd or registry!"
        Write-Host "MdmUrl            : $($fields.MdmUrl)"
        Write-Host "MdmTouUrl         : $($fields.MdmTouUrl)"
        Write-Host "MdmComplianceUrl  : $($fields.MdmComplianceUrl)"
        Write-Host "AzureAdPrtAuthority: $($fields.AzureAdPrtAuthority)"

        if (-not $fields.MdmUrl) {
            Write-Warning "MdmUrl is blank; this can cause enrollment failures."
        }
        if (-not $fields.MdmTouUrl) {
            Write-Warning "MdmTouUrl is blank; this can cause enrollment failures."
        }
        if (-not $fields.MdmComplianceUrl) {
            Write-Warning "MdmComplianceUrl is blank; this can cause enrollment failures."
        }
        Write-Host "Sometimes these values take a bit to populate; verify in the Intune portal."

        if ($fields.AzureAdPrtAuthority -like "*common/UserRealm*") {
            Write-Warning "AzureAdPrtAuthority shows a known fallback/error URL: $($fields.AzureAdPrtAuthority)"
        }
    }
}

###############################################################################
# 5. Determine final join state & update NinjaOne custom field
###############################################################################
Write-Host "=== 5. Determine final join state & update NinjaOne custom field ==="
# Re-check all final states
$finalDsregOutput = dsregcmd /status | Out-String
$fields = Get-DsregStatusFields -DsregOutput $finalDsregOutput
$finalDsregMDM = $fields.MDMEnrolled
$finalRegMDM   = Test-LocalIntuneEnrollmentRegistry
$finalMDMEnrolled = ($finalDsregMDM -or $finalRegMDM)

$joinState = "None"
if ($fields.AzureAdJoined -and $finalMDMEnrolled) {
    $joinState = "Both"
}
elseif ($fields.AzureAdJoined) {
    $joinState = "Entra"
}
elseif ($finalMDMEnrolled) {
    $joinState = "Intune"
}

Write-Host "Final device join state: $joinState"

try {
    Ninja-Property-Set entraIntuneJoinState $joinState
    Write-Host "Updated NinjaOne custom field 'entraIntuneJoinState' to '$joinState'."
}
catch {
    Write-Warning "Could not set 'entraIntuneJoinState': $($_.Exception.Message)"
}

Write-Host "Script complete."
