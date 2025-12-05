<#  
.SYNOPSIS  
    Toggle NTLM usage (block vs allow) by editing the three “Restrict NTLM” registry values.  
.DESCRIPTION  
    Mode 0 = Allow NTLM (remove restrictions)  
    Mode 1 = Block NTLM entirely (deny all)  

    Example in NinjaOne:  
        Parameter 1 → 1   # block  
        Parameter 1 → 0   # allow  
#>

param(
    [ValidateSet('0','1')]
    [string]$Mode
)

# registry targets
$keys = @(
    'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0',
    'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
)

# value map: 2 = Deny all, 0 = Allow all
$desired = if ($Mode -eq '1') { 2 } else { 0 }

Write-Host "Setting NTLM policy value to $desired"

# helper
function Set-Dword {
    param($Path, $Name, [int]$Value)
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord
}

# block or allow
Set-Dword "$($keys[0])" 'RestrictReceivingNTLMTraffic' $desired
Set-Dword "$($keys[1])" 'RestrictSendingNTLMTraffic'  $desired
Set-Dword "$($keys[1])" 'RestrictNTLMInDomain'        $desired

# optional: enforce NTLMv2 only when blocking
if ($Mode -eq '1') {
    Set-Dword 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel' 5  # send NTLMv2 only, refuse LM/NTLMv1 :contentReference[oaicite:1]{index=1}
}

# flip same values under Policies hive to try and out‑rank a local cached GPO
foreach ($polPath in @(
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows\NTLM',       # older OSes
    'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'  # redundancy
)) {
    if (-not (Test-Path $polPath)) { New-Item -Path $polPath -Force | Out-Null }
    Set-Dword $polPath 'RestrictReceivingNTLMTraffic' $desired
    Set-Dword $polPath 'RestrictSendingNTLMTraffic'  $desired
    Set-Dword $polPath 'RestrictNTLMInDomain'        $desired
}

Write-Host "NTLM policy updated. No reboot needed, but gpupdate /force is recommended."
exit 0
