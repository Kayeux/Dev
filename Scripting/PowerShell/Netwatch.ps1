param(
  [int]$IntervalSec = 5,
  [string]$LogDir = "C:\ProgramData\NetWatch",
  [string]$PingTarget = "1.1.1.1"
)

New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$log = Join-Path $LogDir ("netwatch_{0}.csv" -f (Get-Date -Format "yyyyMMdd"))

function Get-EffectiveDefaultRoute {
  $if4 = Get-NetIPInterface -AddressFamily IPv4 -ErrorAction SilentlyContinue
  $routes = Get-NetRoute -AddressFamily IPv4 -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue
  if (-not $routes) { return $null }

  $best = $routes | ForEach-Object {
    $iface = $if4 | Where-Object { $_.InterfaceIndex -eq $_.InterfaceIndex } | Select-Object -First 1
    $ifaceMetric = if ($iface) { [int]$iface.InterfaceMetric } else { 9999 }
    [pscustomobject]@{
      InterfaceIndex = $_.InterfaceIndex
      NextHop        = $_.NextHop
      RouteMetric    = [int]$_.RouteMetric
      IfMetric       = $ifaceMetric
      EffectiveMetric= ([int]$_.RouteMetric + $ifaceMetric)
    }
  } | Sort-Object EffectiveMetric, RouteMetric | Select-Object -First 1

  $alias = (Get-NetAdapter -InterfaceIndex $best.InterfaceIndex -ErrorAction SilentlyContinue).Name
  $best | Add-Member -NotePropertyName InterfaceAlias -NotePropertyValue $alias -Force
  return $best
}

function Get-RttMs {
  try {
    $p = New-Object System.Net.NetworkInformation.Ping
    $r = $p.Send($PingTarget, 1000)
    if ($r.Status -eq "Success") { return [int]$r.RoundtripTime }
    return $null
  } catch { return $null }
}

if (-not (Test-Path $log)) {
  "Time,IPv4Connectivity,ActiveDefaultIF,NextHop,EffectiveMetric,EthStatus,EthLink,WifiStatus,WifiLink,PingMs,EthRxDiscard,EthTxDiscard,WifiRxDiscard,WifiTxDiscard" |
    Out-File -FilePath $log -Encoding ASCII
}

while ($true) {
  $t = Get-Date
  $prof = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Select-Object -First 1
  $ipv4Conn = if ($prof) { $prof.IPv4Connectivity } else { "" }

  $def = Get-EffectiveDefaultRoute
  $activeIf = if ($def) { $def.InterfaceAlias } else { "" }
  $nextHop  = if ($def) { $def.NextHop } else { "" }
  $effMet   = if ($def) { $def.EffectiveMetric } else { "" }

  $eth  = Get-NetAdapter -Name "Ethernet" -ErrorAction SilentlyContinue
  $wifi = Get-NetAdapter -Name "Wi-Fi"   -ErrorAction SilentlyContinue

  $ethStats  = if ($eth)  { Get-NetAdapterStatistics -Name $eth.Name  -ErrorAction SilentlyContinue } else { $null }
  $wifiStats = if ($wifi) { Get-NetAdapterStatistics -Name $wifi.Name -ErrorAction SilentlyContinue } else { $null }

  $pingMs = Get-RttMs

  $line = "{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}" -f `
    ($t.ToString("s")), `
    $ipv4Conn, `
    $activeIf, `
    $nextHop, `
    $effMet, `
    ($eth.Status), ($eth.LinkSpeed), `
    ($wifi.Status), ($wifi.LinkSpeed), `
    ($pingMs), `
    ($ethStats.ReceivedDiscardedPackets), ($ethStats.OutboundDiscardedPackets), `
    ($wifiStats.ReceivedDiscardedPackets), ($wifiStats.OutboundDiscardedPackets)

  Add-Content -Path $log -Value $line -Encoding ASCII
  Start-Sleep -Seconds $IntervalSec
}
