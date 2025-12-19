# Close common M365 processes
$procNames = @("outlook","teams","ms-teams","onedrive","winword","excel","powerpnt","msedgewebview2")
Get-Process -ErrorAction SilentlyContinue | Where-Object { $procNames -contains $_.Name } | Stop-Process -Force -ErrorAction SilentlyContinue

# Clear WAM / AAD Broker caches
$paths = @(
  "$env:LOCALAPPDATA\Microsoft\IdentityCache",
  "$env:LOCALAPPDATA\Microsoft\OneAuth",
  "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker",
  "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\LocalState"
)

foreach ($p in $paths) {
  if (Test-Path $p) {
    Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
  }
}

Write-Host "Done."
