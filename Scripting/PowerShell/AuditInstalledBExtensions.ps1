# Enumerate Chrome and Edge extensions for all local user profiles.
# Native PowerShell only. Outputs a clean object list.

$ErrorActionPreference = 'SilentlyContinue'

function Resolve-ExtensionName {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ManifestPath
    )

    $name = $null
    $nameRaw = $null

    try {
        $manifestJson = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $nameRaw = [string]$manifestJson.name

        # Try to resolve __MSG_*__ names
        if ($nameRaw -match '^__MSG_(.+)__$') {
            $key = $Matches[1]
            $defaultLocale = [string]$manifestJson.default_locale

            if ($defaultLocale) {
                $manifestDir = Split-Path -Parent $ManifestPath
                $messagesPath = Join-Path $manifestDir ("_locales\" + $defaultLocale + "\messages.json")

                if (Test-Path -LiteralPath $messagesPath) {
                    $messagesJson = Get-Content -LiteralPath $messagesPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($messagesJson.$key -and $messagesJson.$key.message) {
                        $name = [string]$messagesJson.$key.message
                    }
                }
            }
        }

        if (-not $name) { $name = $nameRaw }
        return @{ Name = $name; NameRaw = $nameRaw; Version = [string]$manifestJson.version }
    }
    catch {
        return @{ Name = $null; NameRaw = $null; Version = $null }
    }
}

function Get-BrowserExtensions {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Chrome","Edge")]
        [string]$Browser
    )

    $results = New-Object System.Collections.Generic.List[object]

    $userRoots = Get-ChildItem -Path "C:\Users" -Directory |
        Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") }

    foreach ($u in $userRoots) {
        $base = if ($Browser -eq "Chrome") {
            Join-Path $u.FullName "AppData\Local\Google\Chrome\User Data"
        } else {
            Join-Path $u.FullName "AppData\Local\Microsoft\Edge\User Data"
        }

        if (-not (Test-Path -LiteralPath $base)) { continue }

        $profiles = Get-ChildItem -LiteralPath $base -Directory |
            Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" }

        foreach ($p in $profiles) {
            $extRoot = Join-Path $p.FullName "Extensions"
            if (-not (Test-Path -LiteralPath $extRoot)) { continue }

            $idDirs = Get-ChildItem -LiteralPath $extRoot -Directory
            foreach ($idDir in $idDirs) {
                $verDirs = Get-ChildItem -LiteralPath $idDir.FullName -Directory
                if (-not $verDirs) { continue }

                # Pick the highest version folder if parseable
                $latest = $verDirs |
                    Sort-Object {
                        try { [version]$_.Name } catch { [version]"0.0.0.0" }
                    } -Descending |
                    Select-Object -First 1

                $manifestPath = Join-Path $latest.FullName "manifest.json"
                if (-not (Test-Path -LiteralPath $manifestPath)) { continue }

                $nameInfo = Resolve-ExtensionName -ManifestPath $manifestPath

                $results.Add([pscustomobject]@{
                    Browser      = $Browser
                    User         = $u.Name
                    Profile      = $p.Name
                    ExtensionId  = $idDir.Name
                    Name         = $nameInfo.Name
                    NameRaw      = $nameInfo.NameRaw
                    Version      = $nameInfo.Version
                    ManifestPath = $manifestPath
                })
            }
        }
    }

    return $results
}

$chrome = Get-BrowserExtensions -Browser "Chrome"
$edge   = Get-BrowserExtensions -Browser "Edge"

$all = $chrome + $edge

# Show results
$all |
    Sort-Object Browser, User, Profile, Name, ExtensionId |
    Format-Table -AutoSize

# Optional: export for allowlist building
# $all | Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$env:PUBLIC\browser_extensions_inventory.csv"