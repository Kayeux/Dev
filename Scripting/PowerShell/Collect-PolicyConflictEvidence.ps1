<#
Collect-PolicyConflictEvidence.ps1
Purpose: Collect evidence for Intune "Conflict" without eyeballing GPO vs Intune.
No external modules. Uses gpresult + MdmDiagnosticsTool + expand.exe.

Usage:
  .\Collect-PolicyConflictEvidence.ps1
  .\Collect-PolicyConflictEvidence.ps1 -SearchTerms "MDMWinsOverGP","Password","BitLocker","./Device/Vendor/MSFT/..."
#>

[CmdletBinding()]
param(
  [string[]]$SearchTerms = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-Folder([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outRoot = Join-Path $env:TEMP "PolicyConflict_$ts"
New-Folder $outRoot

"Output: $outRoot" | Tee-Object -FilePath (Join-Path $outRoot "README.txt") | Out-Null

# --- 1) gpresult (built-in) ---
$gpComputerHtml = Join-Path $outRoot "gpresult_computer.html"
$gpUserHtml     = Join-Path $outRoot "gpresult_user.html"
$gpComputerXml  = Join-Path $outRoot "gpresult_computer.xml"
$gpUserXml      = Join-Path $outRoot "gpresult_user.xml"

& gpresult /scope computer /h $gpComputerHtml /f | Out-Null
& gpresult /scope user     /h $gpUserHtml     /f | Out-Null
& gpresult /scope computer /x $gpComputerXml  /f | Out-Null
& gpresult /scope user     /x $gpUserXml      /f | Out-Null

# --- 2) MDM Diagnostics (built-in on Win10/11) ---
$mdmTool = Join-Path $env:WINDIR "System32\MdmDiagnosticsTool.exe"
$mdmCab  = Join-Path $outRoot "MDMDiagReport.cab"
$mdmOut  = Join-Path $outRoot "MDMDiagnostics_out"
$mdmExtract = Join-Path $outRoot "MDMDiagExtracted"
New-Folder $mdmExtract

$usedCab = $false
if (Test-Path -LiteralPath $mdmTool) {
  try {
    # Try targeted areas first (faster/smaller)
    $areas = "DeviceEnrollment;DeviceProvisioning;Policy"
    $p = Start-Process -FilePath $mdmTool -ArgumentList @("-area",$areas,"-cab",$mdmCab) -Wait -PassThru -WindowStyle Hidden
    if ($p.ExitCode -eq 0 -and (Test-Path -LiteralPath $mdmCab)) { $usedCab = $true }
  } catch {
    $usedCab = $false
  }

  if (-not $usedCab) {
    # Fallback: dump to folder (works on more builds)
    New-Folder $mdmOut
    Start-Process -FilePath $mdmTool -ArgumentList @("-out",$mdmOut) -Wait -WindowStyle Hidden | Out-Null
  }
} else {
  "WARNING: MdmDiagnosticsTool.exe not found at $mdmTool" | Add-Content (Join-Path $outRoot "README.txt")
}

if ($usedCab) {
  # expand.exe is built-in
  & expand.exe -F:* $mdmCab $mdmExtract | Out-Null
}

# --- 3) Build a quick "what’s blocking GP" + keyword hit list ---
$summary = Join-Path $outRoot "SUMMARY_hits.txt"
$needles = @(
  "MDM Uris Blocking GP",
  "Blocking GP",
  "Blocked GP",
  "MDMWinsOverGP",
  "ControlPolicyConflict"
) + $SearchTerms

$targets = @(
  $gpComputerXml, $gpUserXml,
  (Join-Path $mdmExtract "MDMDiagReport.xml"),
  (Join-Path $mdmExtract "MDMDiagHtmlReport.html")
)

# Also scan any extracted *.xml/*.html we can find (without going insane)
if (Test-Path -LiteralPath $mdmExtract) {
  $targets += Get-ChildItem -Path $mdmExtract -Recurse -File -Include *.xml,*.html -ErrorAction SilentlyContinue |
             Select-Object -ExpandProperty FullName
}
$targets = $targets | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -Unique

"=== Targets scanned ===" | Out-File $summary -Encoding ASCII
$targets | Out-File $summary -Append -Encoding ASCII
"`n=== Hits ===" | Out-File $summary -Append -Encoding ASCII

foreach ($t in $targets) {
  foreach ($n in ($needles | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique)) {
    $hits = Select-String -Path $t -SimpleMatch -Pattern $n -ErrorAction SilentlyContinue
    if ($hits) {
      "---- $n in $t ----" | Out-File $summary -Append -Encoding ASCII
      $hits | ForEach-Object { $_.Line } | Out-File $summary -Append -Encoding ASCII
      "" | Out-File $summary -Append -Encoding ASCII
    }
  }
}

"Done. Review: $summary" | Tee-Object -FilePath (Join-Path $outRoot "README.txt") -Append | Out-Null
"Open gpresult HTMLs if needed:`n  $gpComputerHtml`n  $gpUserHtml" | Add-Content (Join-Path $outRoot "README.txt")
if ($usedCab) { "MDM CAB extracted to: $mdmExtract" | Add-Content (Join-Path $outRoot "README.txt") }