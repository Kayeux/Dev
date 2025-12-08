<#
Browser Extension Snapshot -> Datto RMM UDF7 (Custom7)
Version: 2.1

Changes vs v2.0:
- Removes single-line delimiter mode (no pipe-joined output).
- Adds hard fallback: if UDF payload > 255 chars, write summary:
  "\<\#\> Extensions \<\#\> Browsers"
- Still honors:
  - one UDF only (Custom7)
  - sorted output
  - IDs when <= 8 total extensions
  - compressed names when > 8
  - Firefox support
  - local log path under ProgramData\ScriptWork
- Uses REG ADD for reliable agent pickup.

Run as SYSTEM in Datto.
#>

param(
    [string]$InlineText = ""
)

$ErrorActionPreference = "SilentlyContinue"

# ----------------------------
# Config
# ----------------------------
$UdfValueName = "Custom7"
$RegKeyPath   = "HKLM\SOFTWARE\CentraStage"

$LogDir  = "C:\ProgramData\ScriptWork\ExtensionAudit"
$LogFile = Join-Path $LogDir "browser_extensions_udf7.txt"

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

        # Resolve __MSG_*__ names if possible
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
            } catch {}

            $results.Add([pscustomobject]@{
                ExtensionId = $id
                Name        = $nm
            })
        }
    }
    catch { }

    return $results
}

function Compress-Name {
    param([string]$Name, [string]$FallbackId)

    $n = $Name
    if ([string]::IsNullOrWhiteSpace($n)) { $n = $FallbackId }

    # Remove whitespace to save chars
    $n = ($n -replace "\s+","").Trim()

    # Cap to 32 chars
    if ($n.Length -gt 32) { $n = $n.Substring(0,32) }

    return $n
}

function Build-Multiline {
    param(
        [string[]]$Items,
        [string]$InlineText
    )

    $itemsSorted = $Items |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object

    $body = ($itemsSorted -join "`r`n")
    if ($InlineText) { return ($InlineText + "`r`n" + $body) }
    return $body
}

function Get-RegExePath {
    if (Test-Path "$env:windir\sysnative\reg.exe") {
        return "$env:windir\sysnative\reg.exe"
    }
    return "$env:windir\system32\reg.exe"
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
# Build unique device-wide set with browser context
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
    }

$totalExtensions = ($unique | Measure-Object).Count
$browserCount = ($records | Select-Object -ExpandProperty Browser | Sort-Object -Unique | Measure-Object).Count

# ----------------------------
# Choose content (IDs vs Names)
# ----------------------------
if ($totalExtensions -gt 8) {
    $items = $unique | ForEach-Object { Compress-Name -Name $_.Name -FallbackId $_.ExtensionId }
} else {
    $items = $unique | Select-Object -ExpandProperty ExtensionId
}

$contentCandidate = Build-Multiline -Items $items -InlineText $InlineText

# ----------------------------
# Hard fallback if >255 chars
# ----------------------------
$summary = "$totalExtensions Extensions, $browserCount Browsers"

if ($contentCandidate.Length -gt 255) {
    # Try to include InlineText if it still fits
    if ($InlineText) {
        $summaryWithInline = Build-Multiline -Items @($summary) -InlineText $InlineText
        if ($summaryWithInline.Length -le 255) {
            $content = $summaryWithInline
        } else {
            $content = $summary
        }
    } else {
        $content = $summary
    }
}
else {
    $content = $contentCandidate
}

# ----------------------------
# Local log
# ----------------------------
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

$ts = (Get-Date).ToString("s")
$ctx = [Environment]::UserName

$logLines = @()
$logLines += "Version=2.1"
$logLines += "Timestamp=$ts"
$logLines += "Context=$ctx"
$logLines += "TotalUniqueExtensions=$totalExtensions"
$logLines += "BrowserCount=$browserCount"
if ($InlineText) { $logLines += "InlineText=$InlineText" }
$logLines += "UDF7PayloadLength=$($content.Length)"
$logLines += "UDF7Payload=$content"
$logLines += "---- Unique Extensions (ID | Name | Browsers) ----"

$unique |
    Sort-Object ExtensionId |
    ForEach-Object {
        $nm = $_.Name
        if ([string]::IsNullOrWhiteSpace($nm)) { $nm = "" }
        $logLines += ($_.ExtensionId + " | " + $nm + " | " + $_.Browsers)
    }

$logLines | Set-Content -LiteralPath $LogFile -Encoding UTF8

# ----------------------------
# Populate UDF7 via REG ADD (ONLY Custom7)
# ----------------------------
$regExe = Get-RegExePath

& $regExe ADD "$RegKeyPath" /f | Out-Null
& $regExe ADD "$RegKeyPath" /v $UdfValueName /t REG_SZ /d "$content" /f | Out-Null

Write-Host "Wrote ONLY Custom7 via REG ADD."
Write-Host "Total unique extensions detected: $totalExtensions"
Write-Host "Browsers detected: $browserCount"
Write-Host "UDF7 payload length: $($content.Length)"
Write-Host "Local log: $LogFile"