<#
.SYNOPSIS
  Clears Microsoft Teams cache (classic + new work/school) for one or more user profiles.
  Safe to run as the logged-on user OR as SYSTEM.

.NOTES
  - Only deletes cache/temp locations.
  - Does NOT remove Teams binaries or core app data.
  - Users may be prompted to sign in to Teams again.
#>

param(
    [switch]$WhatIfMode
)

# Simple logger
$logRoot = 'C:\ProgramData\Aspire'
$logPath = Join-Path $logRoot 'TeamsCacheCleanup.log'
if (-not (Test-Path $logRoot)) {
    New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp`t$Message"
    Add-Content -Path $logPath -Value $line
}

Write-Log "=== Teams cache cleanup starting. User=$($env:USERNAME), WhatIf=$WhatIfMode ==="

# Detect if we are SYSTEM (worst-case DEM execution)
$runningAsSystem = $false
try {
    if ($env:USERNAME -eq 'SYSTEM' -or
        ([Security.Principal.WindowsIdentity]::GetCurrent().Name -like '*SYSTEM')) {
        $runningAsSystem = $true
    }
} catch {
    # ignore
}

# Kill Teams processes globally (classic + new)
$procNames = @('ms-teams', 'teams', 'MSTeams', 'MSTeamsFrontend', 'ms-teams-uwp')
foreach ($name in $procNames) {
    try {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
        if ($procs) {
            Write-Log "Attempting to stop process '$name' (count=$($procs.Count))"
            if (-not $WhatIfMode) {
                $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Log "Error stopping process '$name': $($_.Exception.Message)"
    }
}

function Get-ProfileTargets {
    param([bool]$IsSystem)

    if (-not $IsSystem) {
        # Current user only
        $profilePath = $env:USERPROFILE
        if ($profilePath -and (Test-Path $profilePath)) {
            $appData = $env:APPDATA
            $localApp = $env:LOCALAPPDATA
            [PSCustomObject]@{
                UserName     = $env:USERNAME
                ProfilePath  = $profilePath
                AppData      = $appData
                LocalAppData = $localApp
            }
        }
        return
    }

    # SYSTEM: iterate local profiles under C:\Users (excluding built-ins)
    $root = 'C:\Users'
    if (-not (Test-Path $root)) { return }

    Get-ChildItem $root -Directory | Where-Object {
        $_.Name -notin @('Public','Default','Default User','All Users') -and
        -not ($_.Attributes.ToString() -match 'ReparsePoint')
    } | ForEach-Object {
        $profilePath = $_.FullName
        $appDataPath = Join-Path $profilePath 'AppData\Roaming'
        $localPath   = Join-Path $profilePath 'AppData\Local'

        [PSCustomObject]@{
            UserName     = $_.Name
            ProfilePath  = $profilePath
            AppData      = $appDataPath
            LocalAppData = $localPath
        }
    }
}

# Enumerate target profiles
$targets = Get-ProfileTargets -IsSystem:$runningAsSystem

if (-not $targets) {
    Write-Log "No target profiles found. Exiting."
    return
}

foreach ($t in $targets) {
    Write-Log "Processing profile UserName=$($t.UserName) Path=$($t.ProfilePath)"

    # Classic Teams cache paths
    $classicRoot = Join-Path $t.AppData 'Microsoft\Teams'
    $classicSubFolders = @(
        'blob_storage',
        'Cache',
        'databases',
        'GPUCache',
        'IndexedDB',
        'Local Storage',
        'tmp'
    )

    foreach ($sub in $classicSubFolders) {
        $path = Join-Path $classicRoot $sub
        if (Test-Path -LiteralPath $path) {
            Write-Log "  Clearing classic Teams cache path: $path"
            if (-not $WhatIfMode) {
                try {
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                } catch {
                    Write-Log "    Error clearing $path : $($_.Exception.Message)"
                }
            }
        } else {
            Write-Log "  Path not found (classic): $path"
        }
    }

    # New Teams cache paths (MSIX work/school)
    $mstRoot = Join-Path $t.LocalAppData 'Packages\MSTeams_8wekyb3d8bbwe'

    # FIXED: no stray comma; wrap Join-Path calls in parentheses
    $newCachePaths = @(
        (Join-Path $mstRoot 'LocalCache'),
        (Join-Path $mstRoot 'TempState')
    )

    foreach ($path in $newCachePaths) {
        if (Test-Path -LiteralPath $path) {
            Write-Log "  Clearing new Teams cache path: $path"
            if (-not $WhatIfMode) {
                try {
                    # Clear contents, not the folder itself
                    Remove-Item -LiteralPath (Join-Path $path '*') -Recurse -Force -ErrorAction Stop
                } catch {
                    Write-Log "    Error clearing contents of $path : $($_.Exception.Message)"
                }
            }
        } else {
            Write-Log "  Path not found (new): $path"
        }
    }
}

Write-Log "=== Teams cache cleanup completed. ==="
