# Datto RMM - Browser Extension Snapshot -> UDF7
# Writes to: HKLM:\SOFTWARE\CentraStage\Custom7
# UDFs show only when populated; UDF value limit is 255 chars.
# This script uses ONE UDF only, per your requirement.

$ErrorActionPreference = "SilentlyContinue"

# ----------------------------
# Config
# ----------------------------
$UdfRegistryName = "Custom7"
$RegPath = "HKLM:\SOFTWARE\CentraStage"

$LogDir  = "C:\ProgramData\ScriptWork\ExtensionAudit"
$LogFile = Join-Path $LogDir "browser_extensions_udf7.txt"

# Known Chromium-family user data roots (relative to C:\Users\<user>\)
$BrowserDefs = @(
    @{ Browser = "Chrome";  RelPath = "AppData\Local\Google\Chrome\User Data" },
    @{ Browser = "Edge";    RelPath = "AppData\Local\Microsoft\Edge\User Data" },
    @{ Browser = "Brave";   RelPath = "AppData\Local\BraveSoftware\Brave-Browser\User Data" },
    @{ Browser = "Vivaldi"; RelPath = "AppData\Local\Vivaldi\User Data" }
)

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

    # Try semantic sort; fall back to name sort
    $sorted = $verDirs | Sort-Object {
        try { [version]($_.Name -replace "[^\d\.].*$","") } catch { [version]"0.0.0.0" }
    } -Descending

    return $sorted | Select-Object -First 1
}

function Get-ManifestName {
    param([string]$ManifestPath)

    try {
        $m = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $name = [string]$m.name
        $ver  = [string]$m.version
        return @{ Name = $name; Version = $ver }
    }
    catch {
        return @{ Name = $null; Version = $null }
    }
}

function Compress-Name {
    param([string]$Name, [string]$FallbackId)

    $n = $Name
    if ([string]::IsNullOrWhiteSpace($n)) { $n = $FallbackId }

    # Remove whitespace to save chars
    $n = ($n -replace "\s+","").Trim()

    # Keep it readable-ish but short (<= 32 chars)
    if ($n.Length -gt 32) { $n = $n.Substring(0,32) }

    return $n
}

# ----------------------------
# Collect extension records
# ----------------------------
$records = New-Object System.Collections.Generic.List[object]

foreach ($u in Get-UserDirs) {
    foreach ($b in $BrowserDefs) {
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

                $mi = Get-ManifestName -ManifestPath $manifest

                $records.Add([pscustomobject]@{
                    Browser     = $b.Browser
                    User        = $u.Name
                    Profile     = $p.Name
                    ExtensionId = $idDir.Name
                    Name        = $mi.Name
                    Version     = $mi.Version
                    Manifest    = $manifest
                })
            }
        }
    }
}

# ----------------------------
# Device-level unique set
# ----------------------------
$unique = $records |
    Group-Object ExtensionId |
    ForEach-Object {
        $first = $_.Group | Select-Object -First 1
        [pscustomobject]@{
            ExtensionId = $first.ExtensionId
            Name        = $first.Name
            SeenIn      = ($_.Group.Browser | Sort-Object -Unique) -join ","
        }
    }

# Total unique extensions across all detected browsers
$total = ($unique | Measure-Object).Count

# ----------------------------
# Build content
# ----------------------------
# IDs mode
$idsSorted = $unique.ExtensionId | Sort-Object
$contentIds = ($idsSorted -join "`r`n")

# Decide mode per your rule + safety length check
$useNames = $false
if ($total -gt 8) { $useNames = $true }
if ($contentIds.Length -gt 255) { $useNames = $true }

if ($useNames) {
    $namesList = $unique | ForEach-Object {
        Compress-Name -Name $_.Name -FallbackId $_.ExtensionId
    }

    $namesSorted = $namesList | Sort-Object
    $content = ($namesSorted -join "`r`n")

    # Final guard: hard-trim if still too long
    if ($content.Length -gt 255) {
        # Keep as many lines as fit
        $lines = $namesSorted
        $acc = New-Object System.Collections.Generic.List[string]
        $len = 0
        foreach ($ln in $lines) {
            $addLen = $ln.Length + (if ($acc.Count -gt 0) { 2 } else { 0 }) # CRLF
            if (($len + $addLen) -gt 255) { break }
            $acc.Add($ln)
            $len += $addLen
        }
        $content = ($acc -join "`r`n")
    }
}
else {
    $content = $contentIds
}

# ----------------------------
# Write local log
# ----------------------------
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null

$ts = (Get-Date).ToString("s")
$logLines = @()
$logLines += "Timestamp=$ts"
$logLines += "TotalUniqueExtensions=$total"
$logLines += "Mode=" + (if ($useNames) { "Names" } else { "IDs" })
$logLines += "---- Unique Extensions (ID | Name | Browsers) ----"

$unique |
    Sort-Object ExtensionId |
    ForEach-Object {
        $nm = $_.Name
        if ([string]::IsNullOrWhiteSpace($nm)) { $nm = "" }
        $logLines += ($_.ExtensionId + " | " + $nm + " | " + $_.SeenIn)
    }

$logLines | Set-Content -LiteralPath $LogFile -Encoding UTF8

# ----------------------------
# Populate UDF7 via registry
# ----------------------------
New-Item -Path $RegPath -Force | Out-Null
New-ItemProperty -Path $RegPath -Name $UdfRegistryName -PropertyType String -Value $content -Force | Out-Null

Write-Host "UDF7 populated via $RegPath\$UdfRegistryName"
Write-Host "Total unique extensions detected: $total"
Write-Host "Mode: " + (if ($useNames) { "Names" } else { "IDs" })
Write-Host "Local log: $LogFile"
