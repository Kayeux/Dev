<#
.SYNOPSIS
    Entra ID join and Intune enrollment automation (v2.0-classic, PS-5.1).

.PARAMETER UseUserCredential
    Forces user-credential enrollment and skips SCCM/device logic.

.PARAMETER DryRun
    Shows intended actions without making changes.

.NOTES
    Designed to run under SYSTEM (RMM, NinjaOne, etc.).
    Creates a log under C:\Techneaux\IntuneJoin.
#>

[CmdletBinding()]
param(
    [switch]$UseUserCredential,
    [switch]$DryRun
)

#─────────────────────────────────────────────────────────
#  Constants and logging
#─────────────────────────────────────────────────────────
$LogRoot = 'C:\Techneaux\IntuneJoin'
if (-not (Test-Path $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
$LogFile = Join-Path $LogRoot ("IntuneJoin-{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $line = "[{0:u}] [{1}] {2}" -f (Get-Date),$Level,$Message
    switch ($Level) {
        'WARN' { Write-Warning  $Message }
        'ERR'  { Write-Error    $Message }
        default{ Write-Host     $Message }
    }
    Add-Content -Path $LogFile -Value $line
}

function Bail {
    param([string]$Reason)
    Write-Log $Reason 'ERR'
    try { Ninja-Property-Set entraIntuneJoinState $Reason } catch {}
    exit 1
}

#─────────────────────────────────────────────────────────
#  Utility helpers (PS-5.1 safe)
#─────────────────────────────────────────────────────────
function Parse-Dsreg {
    param([string]$Text)
    $tenant = $null
    if ($Text -match 'TenantId\s*:\s*([0-9a-fA-F-]+)') { $tenant = $Matches[1] }
    [pscustomobject]@{
        AzureAdJoined = ($Text -match 'AzureAdJoined\s*:\s*YES')
        MdmEnrolled   = ($Text -match 'MDM\s*:\s*Microsoft Intune')
        AzureAdPrt    = ($Text -match 'AzureAdPrt\s*:\s*YES')
        TenantId      = $tenant
    }
}

function Test-DomainJoined { (Get-CimInstance Win32_ComputerSystem).PartOfDomain }

function Test-SccmPresent {
    $wmi = Get-CimInstance -Namespace root\ccm -Class SMS_Client -ErrorAction SilentlyContinue
    if ($wmi) { return $true }
    return Test-Path 'HKLM:\SOFTWARE\Microsoft\CCM'
}

function Get-InteractiveUser {
    $u = (Get-CimInstance Win32_ComputerSystem).UserName
    if ($u -and $u -notmatch '^(NT AUTHORITY|SYSTEM|Administrator)$') { return $u }
    return $null
}

#─────────────────────────────────────────────────────────
#  Legacy helpers from the original script (unchanged logic)
#─────────────────────────────────────────────────────────
function Remove-EnterpriseMgmtScheduledTasks {
    Write-Host 'Searching for existing EnterpriseMgmt tasks...'
    $taskPath = '\Microsoft\Windows\EnterpriseMgmt'
    try {
        $tasks = Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue
        if ($tasks) {
            Write-Host "Found $($tasks.Count) task(s) under $taskPath. Removing..."
            foreach ($t in $tasks) {
                try { Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $taskPath -Confirm:$false }
                catch { Write-Warning "Could not remove task $($t.TaskName): $($_.Exception.Message)" }
            }
        } else { Write-Host "No tasks under $taskPath." }
    } catch { Write-Warning "Could not enumerate $taskPath : $($_.Exception.Message)" }
}

function Ensure-MdmRegistryKeys {
    param([string]$TenantId)
    if (-not $TenantId) { Write-Log 'TenantId missing; cannot verify CloudDomainJoin keys' 'WARN'; return }
    $base = "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\$TenantId"
    if (-not (Test-Path $base)) { Write-Log "$base not found" 'WARN'; return }
    $defaults = @{
        MdmEnrollmentUrl = 'https://enrollment.manage.microsoft.com/enrollmentserver/discovery.svc'
        MdmTermsOfUseUrl = 'https://portal.manage.microsoft.com/TermsofUse.aspx'
        MdmComplianceUrl = 'https://portal.manage.microsoft.com/?portalAction=Compliance'
    }
    foreach ($k in $defaults.Keys) {
        $cur = (Get-ItemProperty -Path $base -Name $k -ErrorAction SilentlyContinue).$k
        if (-not $cur) {
            Write-Log "Writing default $k"
            if (-not $DryRun) { Set-ItemProperty -Path $base -Name $k -Value $defaults[$k] -Type String }
        }
    }
}

function Test-IntuneByRegistry {
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path $root)) { return $false }
    Get-ChildItem -Path $root | ForEach-Object {
        $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        if ($p.EnrollmentState -eq 1 -and $p.ProviderID -match '(MS DM Server|Microsoft Device Management)') { return $true }
    }
    return $false
}

function Clear-EnrollmentKeys {
    $root = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $root) {
        Write-Log 'Clearing enrollment registry keys'
        if (-not $DryRun) { Remove-Item "$root\*" -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-DeviceEnroller {
    <#
        .SYNOPSIS
            Runs deviceenroller.exe and returns the raw console output.
        .PARAMETER Mode
            'device'  -> /c /AutoEnrollMDMUsingAADDeviceCredential
            'user'    -> /c /AutoEnrollMDMUsingAADUserCredential
    #>
    param(
        [ValidateSet('device','user')]
        [string]$Mode
    )

    $exe  = "$env:SystemRoot\System32\deviceenroller.exe"
    $args = if ($Mode -eq 'device') {
                '/c','/AutoEnrollMDMUsingAADDeviceCredential'
            } else {
                '/c','/AutoEnrollMDMUsingAADUserCredential'
            }

    Write-Log "Running: $exe $($args -join ' ')"

    if ($DryRun) { return '' }

    # Capture stdout+stderr even if the exe writes to Console API instead of pipes
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $exe
    $psi.Arguments              = $args -join ' '
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $p = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $p.StandardOutput.ReadToEnd()
    $stdErr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    $full = ($stdOut + $stdErr).TrimEnd()

    if ($full) {
        Write-Host $full                    # live on console
        Write-Log  "deviceenroller output:`n$full"
    } else {
        Write-Log 'deviceenroller produced no console output'
    }

    return $full
}


function Invoke-KlistRecovery {
    Write-Log 'Running klist purge and get'
    if (-not $DryRun) {
        klist purge -li 0x3e7 | Out-Null
        $d = (Get-CimInstance Win32_ComputerSystem).Domain.ToLower()
        klist get "azuread_krbtgt/$d" -li 0x3e7 | Out-Null
    }
}

#─────────────────────────────────────────────────────────
#  New dual-context PRT refresher
#─────────────────────────────────────────────────────────
function Refresh-PRT {
    param([string]$InteractiveUser)

    if ($InteractiveUser) {
        Write-Log ("UserExec : {0}" -f $InteractiveUser)
        if (-not $DryRun) {
            $task = "IntuneJoinUserPRT_$([guid]::NewGuid().Guid)"
            $time = (Get-Date).AddMinutes(1).ToString('HH:mm')
            schtasks /Create /TN $task /TR "`"$env:SystemRoot\System32\dsregcmd.exe /refreshprt`"" `
                    /SC ONCE /ST $time /RU $InteractiveUser /IT /F | Out-Null
            schtasks /Run /TN $task | Out-Null
            Start-Sleep 65
            schtasks /Delete /TN $task /F | Out-Null
        }
    } else {
        Write-Verbose 'No interactive user session found - skipping user-context refresh'
    }

    Write-Verbose 'Refreshing PRT in SYSTEM context'
    if (-not $DryRun) { & "$env:SystemRoot\System32\dsregcmd.exe" /refreshprt | Out-Null }
}

#─────────────────────────────────────────────────────────
#  Event-log watcher fix
#─────────────────────────────────────────────────────────
function Wait-EnrollmentEvent {
    param([int]$Minutes = 5)

    $src = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider'
    $ids = 2900,2914
    $since = (Get-Date).AddMinutes(-$Minutes)
    $f = @{ LogName = $src; Id = $ids; StartTime = $since }

    try {
        $ev = Get-WinEvent -FilterHashtable $f -MaxEvents 25 -ErrorAction Stop
        return ($ev.Count -gt 0)
    } catch {
        $ms = $Minutes * 60000
        $xp = "*[System[(EventID=2900 or EventID=2914) and TimeCreated[timediff(@SystemTime) <= $ms]]]"
        $ev = Get-WinEvent -LogName $src -FilterXPath $xp -MaxEvents 25 -ErrorAction SilentlyContinue
        return ($null -ne $ev )
    }
}

#─────────────────────────────────────────────────────────
#  Script start
#─────────────────────────────────────────────────────────
Write-Log '======== Intune Join Script v2.0 ========'

# Domain prereq
if (-not (Test-DomainJoined)) { Bail 'Machine is not domain joined - exiting' }

# Choose enrollment mode
$Sccm = Test-SccmPresent
$EnrollMode = if ($UseUserCredential) { 'user' } elseif ($Sccm) { 'device' } else { 'user' }
Write-Log "SCCM detected: $Sccm  |  Enrollment mode: $EnrollMode"

# Remove legacy auto-enroll tasks
if (-not $DryRun) { Remove-EnterpriseMgmtScheduledTasks }

# Initial dsreg snapshot
$dsText = dsregcmd /status | Out-String
$fields = Parse-Dsreg $dsText

# Entra join if needed
if (-not $fields.AzureAdJoined) {
    Write-Log 'Device not Entra joined - starting join'
    Clear-EnrollmentKeys
    if (-not $DryRun) { & "$env:SystemRoot\System32\dsregcmd.exe" /debug /join | Out-Null }
    Start-Sleep 15
    $fields = Parse-Dsreg (dsregcmd /status | Out-String)
    if (-not $fields.AzureAdJoined) { Bail 'Entra join failed' }
} else {
    Write-Verbose 'Already Entra joined'
}

# Ensure CloudDomainJoin URLs
Ensure-MdmRegistryKeys -TenantId $fields.TenantId

# Intune enrollment logic
if ($fields.MdmEnrolled -or (Test-IntuneByRegistry)) {
    Write-Log 'Intune already detected - skipping enrollment'
} else {
    Write-Log 'Intune not detected - executing enrollment path'
    $user = Get-InteractiveUser
    Refresh-PRT -InteractiveUser $user

    # Verify PRT; if still missing, attempt Kerberos ticket fix then refresh again
    if (-not (Parse-Dsreg (dsregcmd /status | Out-String)).AzureAdPrt) {
        Invoke-KlistRecovery
        Refresh-PRT -InteractiveUser $user
    }

    Invoke-DeviceEnroller -Mode $EnrollMode
    Start-Sleep 15

    if (-not ((Test-IntuneByRegistry) -or (Parse-Dsreg (dsregcmd /status | Out-String)).MdmEnrolled) `
        -and -not (Wait-EnrollmentEvent)) {
        Write-Log 'WARNING: Intune enrollment still not confirmed' 'WARN'
    }
}

# Final state reporting
$final = Parse-Dsreg (dsregcmd /status | Out-String)
$intune = $final.MdmEnrolled -or (Test-IntuneByRegistry)
$state = switch ("$($final.AzureAdJoined)$intune") {
    'TrueTrue'  { 'Both' }
    'TrueFalse' { 'Entra' }
    'FalseTrue' { 'Intune' }
    default     { 'None' }
}
Write-Log "Final join state = $state"
try { Ninja-Property-Set entraIntuneJoinState $state } catch {}

Write-Log '======== Script complete ========'
