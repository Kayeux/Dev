<#
SharePoint/OneDrive path preflight audit (no modules)
- Uses Microsoft Graph + device code flow (delegated permissions)
- Flags:
  * Decoded URL path length > 400 (SharePoint/OneDrive limit)
  * Any single segment (folder/file name) > 255
  * Optional: predicted local path length > WindowsLimit (default 260)

Refs:
- SPO/OD path limit 400 decoded + 255 segment: https://support.microsoft.com/office/restrictions-and-limitations-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa
- Windows MAX_PATH 260: https://learn.microsoft.com/windows/win32/fileio/maximum-file-path-limitation
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$TenantId,

  [Parameter(Mandatory=$true)]
  [string]$ClientId,

  # Scan ONE site by URL (recommended per-site for speed and clearer output)
  [Parameter(Mandatory=$true)]
  [string]$SiteUrl,

  # SharePoint/OneDrive decoded path limit
  [int]$SharePointLimit = 400,

  # Windows MAX_PATH style threshold (only used if -LocalRoot is provided)
  [int]$WindowsLimit = 260,

  # Optional: local root to estimate full local path length
  # Example: "C:\Users\Kaleb\OneDrive - PumpTech"
  [string]$LocalRoot = "",

  # Output CSV
  [string]$OutCsv = ".\LongPaths.csv",

  # Set to scan only files (skip folder rows in output)
  [switch]$FilesOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Invoke-Graph {
  param(
    [Parameter(Mandatory=$true)][string]$Method,
    [Parameter(Mandatory=$true)][string]$Uri,
    [Parameter()][hashtable]$Headers,
    [Parameter()][object]$Body
  )

  $maxRetries = 6
  $delay = 2

  for ($i=0; $i -le $maxRetries; $i++) {
    try {
      if ($null -ne $Body) {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body -ContentType "application/x-www-form-urlencoded"
      } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers
      }
    } catch {
      $status = $null
      try { $status = $_.Exception.Response.StatusCode.value__ } catch { }

      # 429/503 retry with backoff
      if ($status -in 429, 503) {
        Start-Sleep -Seconds $delay
        $delay = [Math]::Min($delay * 2, 30)
        continue
      }
      throw
    }
  }

  throw "Graph request failed after retries: $Method $Uri"
}

function Read-WebResponseBody {
  param([Parameter(Mandatory=$true)]$ErrorRecord)

  try {
    $resp = $ErrorRecord.Exception.Response
    if ($null -eq $resp) { return $null }
    $stream = $resp.GetResponseStream()
    if ($null -eq $stream) { return $null }
    $reader = New-Object System.IO.StreamReader($stream)
    $text = $reader.ReadToEnd()
    $reader.Close()
    return $text
  } catch {
    return $null
  }
}

function Get-DeviceCodeToken {
  param(
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [string]$Scope = "https://graph.microsoft.com/Sites.Read.All offline_access"
  )

  # WinPS 5.1 sometimes needs this
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

  $dcUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode"
  $dc = Invoke-RestMethod -Method POST -Uri $dcUri -Body @{
    client_id = $ClientId
    scope     = $Scope
  } -ContentType "application/x-www-form-urlencoded"

  # Make it fast for you:
  try { Set-Clipboard -Value $dc.user_code } catch {}
  try { Start-Process $dc.verification_uri } catch {}

  Write-Host ""
  Write-Host ("Device code: {0} (copied to clipboard)" -f $dc.user_code)
  Write-Host ("Open: {0}" -f $dc.verification_uri)
  Write-Host ("You have ~{0} seconds to finish sign-in (default 900 / 15 min)." -f $dc.expires_in)
  Write-Host ""

  $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
  $interval = [int]$dc.interval
  $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
  Start-Sleep 10
  while ((Get-Date) -lt $deadline) {

    Start-Sleep -Seconds $interval

    try {
      $tok = Invoke-RestMethod -Method POST -Uri $tokenUri -Body @{
        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
        client_id   = $ClientId
        device_code = $dc.device_code
      } -ContentType "application/x-www-form-urlencoded"

      return $tok.access_token
    }
    catch {
      $err = Get-JsonErrorBody -ErrorRecord $_
      if ($null -eq $err) { throw }

      switch ($err.error) {
        # EXPECTED while user hasn't finished auth; keep polling.
        "authorization_pending" { continue }  # expected per Microsoft docs :contentReference[oaicite:2]{index=2}
        "slow_down"             { $interval += 5; continue }
        "authorization_declined" { throw "User declined the device login." } # MS docs :contentReference[oaicite:3]{index=3}
        "expired_token"         { throw "Device code expired. Re-run to get a new code." } # MS docs :contentReference[oaicite:4]{index=4}
        default                 { throw ("Device flow failed: {0} | {1}" -f $err.error, $err.error_description) }
      }
    }
  }

  throw "Timed out waiting for device login (device_code expired)."
}

function Get-JsonErrorBody {
  param([Parameter(Mandatory=$true)]$ErrorRecord)

  try {
    $resp = $ErrorRecord.Exception.Response
    if ($null -eq $resp) { return $null }
    $stream = $resp.GetResponseStream()
    if ($null -eq $stream) { return $null }
    $reader = New-Object System.IO.StreamReader($stream)
    $text = $reader.ReadToEnd()
    $reader.Close()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return ($text | ConvertFrom-Json)
  } catch {
    return $null
  }
}


function Get-GraphSite {
  param([string]$SiteUrl, [hashtable]$Headers)

  $u = [Uri]$SiteUrl
  $hostPath = $u.Host
  $path = $u.AbsolutePath  # includes leading "/"
  $siteLookup = "https://graph.microsoft.com/v1.0/sites/"+$hostPath+":"+$path
  return Invoke-Graph -Method GET -Uri $siteLookup -Headers $Headers
}

function Get-AllPages {
  param([string]$FirstUrl, [hashtable]$Headers)

  $url = $FirstUrl
  while ($null -ne $url -and $url -ne "") {
    $resp = Invoke-Graph -Method GET -Uri $url -Headers $Headers
    foreach ($v in $resp.value) { $v }
    $url = $resp.'@odata.nextLink'
  }
}

# Acquire token
$token = Get-DeviceCodeToken -TenantId $TenantId -ClientId $ClientId
$headers = @{ Authorization = "Bearer $token" }

# Resolve site + derive site-relative prefix for "without-site" calculations
$site = Get-GraphSite -SiteUrl $SiteUrl -Headers $headers
$siteWeb = [Uri]$site.webUrl
$siteRel = [Uri]::UnescapeDataString($siteWeb.AbsolutePath).TrimStart("/")  # e.g. "sites/FrostEngineering"

# Prep CSV
if (Test-Path $OutCsv) { Remove-Item $OutCsv -Force }
$buffer = New-Object System.Collections.Generic.List[object]
$flushEvery = 500

function Flush-Buffer {
  if ($buffer.Count -gt 0) {
    $buffer | Export-Csv -Path $OutCsv -NoTypeInformation -Append
    $buffer.Clear() | Out-Null
  }
}

Write-Host "Scanning site: $($site.webUrl)"
Write-Host "SiteRel prefix: $siteRel"
Write-Host ""

# Drives in the site
$drivesUrl = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$select=id,name,webUrl"
$drives = @(Get-AllPages -FirstUrl $drivesUrl -Headers $headers)

foreach ($d in $drives) {
  Write-Host "Drive: $($d.name)"

  # BFS queue of folder item IDs; start from root
  $root = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$($d.id)/root?`$select=id,name" -Headers $headers
  $q = New-Object System.Collections.Generic.Queue[string]
  $q.Enqueue($root.id)

  while ($q.Count -gt 0) {
    $folderId = $q.Dequeue()

    $childrenUrl = "https://graph.microsoft.com/v1.0/drives/$($d.id)/items/$folderId/children?`$top=999&`$select=id,name,folder,file,webUrl"
    foreach ($item in Get-AllPages -FirstUrl $childrenUrl -Headers $headers) {

      if ($null -ne $item.folder) {
        $q.Enqueue($item.id)
      }

      if ($FilesOnly -and ($null -eq $item.file)) { continue }
      if ([string]::IsNullOrWhiteSpace($item.webUrl)) { continue }

      $iu = [Uri]$item.webUrl
      $decodedPath = [Uri]::UnescapeDataString($iu.AbsolutePath).TrimStart("/")  # excludes domain
      $spLen = $decodedPath.Length

      $segments = $decodedPath.Split("/")
      $maxSegLen = 0
      foreach ($s in $segments) {
        if ($s.Length -gt $maxSegLen) { $maxSegLen = $s.Length }
      }

      # Remove the site prefix if it matches (gives you the "tail" that will be appended under a new site URL)
      $tail = $decodedPath
      if ($decodedPath.StartsWith($siteRel + "/")) {
        $tail = $decodedPath.Substring($siteRel.Length + 1)
      } elseif ($decodedPath -eq $siteRel) {
        $tail = ""
      }

      $localLen = ""
      $exceedsLocal = $false
      if (-not [string]::IsNullOrWhiteSpace($LocalRoot)) {
        $localPath = ($LocalRoot.TrimEnd("\") + "\" + ($decodedPath -replace "/","\"))
        $localLen = $localPath.Length
        $exceedsLocal = ($localLen -gt $WindowsLimit)
      }

      $exceeds400 = ($spLen -gt $SharePointLimit)
      $exceeds255 = ($maxSegLen -gt 255)

      if ($exceeds400 -or $exceeds255 -or $exceedsLocal) {
        $buffer.Add([pscustomobject]@{
          SiteUrl                = $site.webUrl
          DriveName              = $d.name
          ItemType               = $(if ($null -ne $item.file) { "File" } elseif ($null -ne $item.folder) { "Folder" } else { "Item" })
          WebUrl                 = $item.webUrl
          DecodedPath            = $decodedPath
          DecodedPathLength      = $spLen
          SiteRelativePrefix     = $siteRel
          TailPathWithoutSite    = $tail
          MaxSegmentLength       = $maxSegLen
          Exceeds400DecodedPath  = $exceeds400
          Exceeds255Segment      = $exceeds255
          LocalPathLength        = $localLen
          ExceedsWindowsLimit    = $exceedsLocal
          ItemId                 = $item.id
        }) | Out-Null

        if ($buffer.Count -ge $flushEvery) { Flush-Buffer }
      }
    }
  }

  Flush-Buffer
}

Write-Host ""
Write-Host "Done. Output: $OutCsv"