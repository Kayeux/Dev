<#
AuditInstalledBrowserExtensions v3.0

Purpose:
- Enumerate extensions across:
  - Chrome
  - Edge
  - Firefox
  - Vivaldi
- Build $content containing full data:
  "ExtensionId | Name | Browsers"
- No char counting
- No UDF writes
- Output full $content to StdOut for Datto email capture
- Write a local log under C:\ProgramData\ScriptWork

Run context:
- Works in user/admin/SYSTEM. SYSTEM is fine for reading most profiles.
#>

param(
    [string]$InlineText = ""
)

$ErrorActionPreference = "SilentlyContinue"

# ----------------------------
# Config
# ----------------------------
$LogDir  = "C:\ProgramData\ScriptWork\ExtensionAudit"
$LogFile = Join-Path $LogDir "browser_extensions_full_stdout.txt"

$ChromiumDefs = @(
    @{ Browser = "Chrome";  RelPath = "AppData\Local\Google\Chrome\User Data" },
    @{ Browser = "Edge";    RelPath = "AppData\Local\Microsoft\Edge\User Data" },
    @{ Browser = "Vivaldi"; RelPath = "AppData\Local\Vivaldi\User Data" }
)

$FirefoxRelRoot = "AppData\Roaming\Mozilla\Firefox\Profiles"

# ----------------------------
# Helpers
# ----------------------------
function Get-UserDirs {
    Get-ChildItem -Path "C:\Users" -Directory |
        Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }
}

function Get-LatestVersionDir {
    param([string]$IdPath)

    $verDirs = Get-ChildItem -LiteralPath $IdPath -Directory
    if (-not $verDirs) { return $null }

    $sorted = $verDirs | Sort-Object {
        try { [version]($_.Name -replace "[^\d\.].*$","") } catch { [version]"0.0.0.0" }
    } -Descending

    return $sorted | Select-Object -First 1
}

function Resolve-ChromiumName {
    param([string]$ManifestPath)

    try {
        $m = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $nameRaw = [string]$m.name

        # Try to resolve __MSG_*__ names
        if ($nameRaw -match '^__MSG_(.+)__$' -and $m.default_locale) {
            $key = $Matches[1]
            $manifestDir = Split-Path -Parent $ManifestPath
            $messagesPath = Join-Path $manifestDir ("_locales\" + $m.default_locale + "\messages.json")

            if (Test-Path -LiteralPath $messagesPath) {
                $msg = Get-Content -LiteralPath $messagesPath -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($msg.$key -and $msg.$key.message) {
                    return [string]$msg.$key.message
                }
            }
        }

        return $nameRaw
    }
    catch {
        return $null
    }
}

function Get-Firefox-ExtensionsFromProfile {
    param([string]$ProfilePath)

    $results = New-Object System.Collections.Generic.List[object]
    $extJson = Join-Path $ProfilePath "extensions.json"

    if (-not (Test-Path -LiteralPath $extJson)) { return $results }

    try {
        $j = Get-Content -LiteralPath $extJson -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $j.addons) { return $results }

        foreach ($a in $j.addons) {
            $type = [string]$a.type
            if ($type -and $type -ne "extension") { continue }

            $id = [string]$a.id
            if ([string]::IsNullOrWhiteSpace($id)) { continue }

            $nm = $null
            try {
                if ($a.defaultLocale -and $a.defaultLocale.name) {
                    $nm = [string]$a.defaultLocale.name
                } elseif ($a.name) {
                    $nm = [string]$a.name
                }
            } catch { }

            $results.Add([pscustomobject]@{
                ExtensionId = $id
                Name        = $nm
            })
        }
    }
    catch { }

    return $results
}

# ----------------------------
# Collect records
# ----------------------------
$records = New-Object System.Collections.Generic.List[object]

foreach ($u in Get-UserDirs) {

    # Chromium browsers
    foreach ($b in $ChromiumDefs) {
        $base = Join-Path $u.FullName $b.RelPath
        if (-not (Test-Path -LiteralPath $base)) { continue }

        $profiles = Get-ChildItem -LiteralPath $base -Directory |
            Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" }

        foreach ($p in $profiles) {
            $extRoot = Join-Path $p.FullName "Extensions"
            if (-not (Test-Path -LiteralPath $extRoot)) { continue }

            $idDirs = Get-ChildItem -LiteralPath $extRoot -Directory
            foreach ($idDir in $idDirs) {
                $latest = Get-LatestVersionDir -IdPath $idDir.FullName
                if (-not $latest) { continue }

                $manifest = Join-Path $latest.FullName "manifest.json"
                if (-not (Test-Path -LiteralPath $manifest)) { continue }

                $nm = Resolve-ChromiumName -ManifestPath $manifest

                $records.Add([pscustomobject]@{
                    Browser     = $b.Browser
                    User        = $u.Name
                    Profile     = $p.Name
                    ExtensionId = $idDir.Name
                    Name        = $nm
                })
            }
        }
    }

    # Firefox
    $ffRoot = Join-Path $u.FullName $FirefoxRelRoot
    if (Test-Path -LiteralPath $ffRoot) {
        $ffProfiles = Get-ChildItem -LiteralPath $ffRoot -Directory
        foreach ($fp in $ffProfiles) {
            $ffExts = Get-Firefox-ExtensionsFromProfile -ProfilePath $fp.FullName
            foreach ($e in $ffExts) {
                $records.Add([pscustomobject]@{
                    Browser     = "Firefox"
                    User        = $u.Name
                    Profile     = $fp.Name
                    ExtensionId = $e.ExtensionId
                    Name        = $e.Name
                })
            }
        }
    }
}

# ----------------------------
# Build unified view: ExtensionId + Name + Browsers
# ----------------------------
$unique = $records |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_.ExtensionId) } |
    Group-Object ExtensionId |
    ForEach-Object {
        $browsers = $_.Group | Select-Object -ExpandProperty Browser | Sort-Object -Unique
        $namePick = $_.Group | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } |
            Select-Object -ExpandProperty Name -First 1

        [pscustomobject]@{
            ExtensionId = $_.Name
            Name        = $namePick
            Browsers    = ($browsers -join ",")
        }
    } |
    Sort-Object ExtensionId

# ----------------------------
# Build $content
# ----------------------------
$lines = foreach ($u in $unique) {
    $nm = $u.Name
    if ([string]::IsNullOrWhiteSpace($nm)) { $nm = "" }
    "$($u.ExtensionId) | $nm | $($u.Browsers)"
}

if ($InlineText) {
    $content = $InlineText + "`r`n" + ($lines -join "`r`n")
} else {
    $content = ($lines -join "`r`n")
}

# ----------------------------
# Local log
# ----------------------------
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

$ts = (Get-Date).ToString("s")
$ctx = [Environment]::UserName

@(
    "Version=3.0"
    "Timestamp=$ts"
    "Context=$ctx"
    "---- ExtensionId | Name | Browsers ----"
    $lines
) | Set-Content -LiteralPath $LogFile -Encoding UTF8

# ----------------------------
# StdOut for Datto email
# ----------------------------
Write-Host $env:COMPUTERNAME
Write-Host $content
Write-Host ""
Write-Host "Local log: $LogFile"
