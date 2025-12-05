<#
.SYNOPSIS
    Hide or restore the ScreenConnect tray icon for ALL users (current and future).
.DESCRIPTION
    - Applies settings to all existing user profiles
    - Sets up default for new users via registry
    - Modifies system.config for system-wide effect
.PARAMETER ShowIcon
    Switch. If supplied, restores the tray icon instead of hiding it.
#>
[CmdletBinding()]
param(
    [Alias('Restore','Unhide')]
    [switch]$ShowIcon
)

$flags = 'ShowSystemTrayIcon','AccessShowSystemTrayIcon','SupportShowSystemTrayIcon'
$desiredVal = if ($ShowIcon.IsPresent) { 'true' } else { 'false' }

Write-Host "Setting ScreenConnect tray icon to '$desiredVal' for ALL users..." -ForegroundColor Cyan

# 1. SYSTEM-WIDE CONFIG (affects all users)
Write-Verbose "Processing system-wide configuration..."
$searchPaths = @(
    'C:\Program Files\ScreenConnect Client*',
    'C:\Program Files (x86)\ScreenConnect Client*',
    'C:\ProgramData\ScreenConnect Client*'
)

$configsModified = 0
foreach ($searchPath in $searchPaths) {
    Get-ChildItem $searchPath -Directory -ErrorAction SilentlyContinue |
    ForEach-Object {
        $cfg = Join-Path $_ 'system.config'
        if (-not (Test-Path $cfg)) { return }
        
        try {
            [xml]$xml = Get-Content $cfg -Raw
            
            # Ensure configuration structure exists
            $sections = $xml.configuration.configSections
            if (-not $sections) {
                $sections = $xml.CreateElement('configSections')
                $xml.configuration.PrependChild($sections) | Out-Null
            }
            
            if (-not $sections.SelectSingleNode("./section[@name='ScreenConnect.UserInterfaceSettings']")){
                $sec = $xml.CreateElement('section')
                $sec.SetAttribute('name','ScreenConnect.UserInterfaceSettings')
                $sec.SetAttribute('type','System.Configuration.ClientSettingsSection')
                $sections.AppendChild($sec) | Out-Null
            }
            
            $ui = $xml.SelectSingleNode('//ScreenConnect.UserInterfaceSettings')
            if (-not $ui) {
                $ui = $xml.CreateElement('ScreenConnect.UserInterfaceSettings')
                $xml.configuration.AppendChild($ui) | Out-Null
            }
            
            # Set the flags
            foreach ($f in $flags) {
                $set = $ui.SelectSingleNode("./setting[@name='$f']")
                if (-not $set) {
                    $set = $xml.CreateElement('setting')
                    $set.SetAttribute('name',$f)
                    $set.SetAttribute('serializeAs','String')
                    $val = $xml.CreateElement('value')
                    $set.AppendChild($val) | Out-Null
                    $ui.AppendChild($set) | Out-Null
                }
                $set.value = $desiredVal
            }
            
            $xml.Save($cfg)
            $configsModified++
            Write-Verbose "Modified system config: $cfg"
            
        } catch {
            Write-Error "Failed to process $cfg`: $($_.Exception.Message)"
        }
    }
}

# 2. CURRENT USER REGISTRY SETTINGS
Write-Verbose "Applying settings to current user registry..."
$regPath = "HKCU:\Software\ScreenConnect"
try {
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    
    foreach ($flag in $flags) {
        Set-ItemProperty -Path $regPath -Name $flag -Value $desiredVal -Type String -Force
    }
    Write-Verbose "Applied registry settings for current user"
} catch {
    Write-Warning "Could not set current user registry: $($_.Exception.Message)"
}

# 3. ALL EXISTING USER PROFILES
Write-Verbose "Applying settings to all existing user profiles..."
$userProfiles = Get-WmiObject -Class Win32_UserProfile | Where-Object { 
    $_.Special -eq $false -and $_.LocalPath -like "C:\Users\*" 
}

foreach ($profile in $userProfiles) {
    $userSID = $profile.SID
    $regPath = "Registry::HKEY_USERS\$userSID\Software\ScreenConnect"
    
    try {
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        foreach ($flag in $flags) {
            Set-ItemProperty -Path $regPath -Name $flag -Value $desiredVal -Type String -Force
        }
        Write-Verbose "Applied settings to user SID: $userSID"
    } catch {
        Write-Verbose "Could not access registry for SID $userSID (may be unloaded)"
    }
}

# 4. DEFAULT USER PROFILE (for new users)
Write-Verbose "Setting defaults for new users..."
$defaultRegPath = "Registry::HKEY_USERS\.DEFAULT\Software\ScreenConnect"
try {
    if (-not (Test-Path $defaultRegPath)) {
        New-Item -Path $defaultRegPath -Force | Out-Null
    }
    
    foreach ($flag in $flags) {
        Set-ItemProperty -Path $defaultRegPath -Name $flag -Value $desiredVal -Type String -Force
    }
    Write-Verbose "Set default profile settings"
} catch {
    Write-Warning "Could not set default user profile: $($_.Exception.Message)"
}

# 5. NTUSER.DAT template modification for future users
Write-Verbose "Modifying default user template..."
$defaultUserPath = "C:\Users\Default\NTUSER.DAT"
if (Test-Path $defaultUserPath) {
    try {
        # Load the default user hive
        $result = Start-Process -FilePath "reg.exe" -ArgumentList "load","HKLM\TEMP_DEFAULT","$defaultUserPath" -Wait -PassThru -WindowStyle Hidden
        
        if ($result.ExitCode -eq 0) {
            $tempRegPath = "Registry::HKEY_LOCAL_MACHINE\TEMP_DEFAULT\Software\ScreenConnect"
            
            if (-not (Test-Path $tempRegPath)) {
                New-Item -Path $tempRegPath -Force | Out-Null
            }
            
            foreach ($flag in $flags) {
                Set-ItemProperty -Path $tempRegPath -Name $flag -Value $desiredVal -Type String -Force
            }
            
            # Unload the hive
            Start-Process -FilePath "reg.exe" -ArgumentList "unload","HKLM\TEMP_DEFAULT" -Wait -WindowStyle Hidden
            Write-Verbose "Modified default user template"
        }
    } catch {
        Write-Warning "Could not modify default user template: $($_.Exception.Message)"
    }
}

# 6. RESTART SERVICES
Write-Verbose "Restarting ScreenConnect services..."
$services = Get-Service 'ScreenConnect Client*' -ErrorAction SilentlyContinue

if ($services) {
    foreach ($service in $services) {
        Stop-Service $service -Force -ErrorAction SilentlyContinue
    }
    
    Start-Sleep -Seconds 3
    Get-Process -Name "*ScreenConnect*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    
    foreach ($service in $services) {
        Start-Service $service -ErrorAction SilentlyContinue
    }
    Write-Verbose "Services restarted"
}

# 7. NOTIFY ALL LOGGED-IN USERS
Write-Verbose "Refreshing system tray for all sessions..."
try {
    # Get all active user sessions
    $sessions = query user 2>$null | Select-Object -Skip 1 | ForEach-Object {
        $parts = $_ -split '\s+', 5
        if ($parts[0] -ne '' -and $parts[1] -match '^\d+$') {
            [PSCustomObject]@{
                Username = $parts[0]
                SessionID = $parts[1]
            }
        }
    }
    
    foreach ($session in $sessions) {
        # Send notification to refresh system tray
        $null = Start-Process -FilePath "msg.exe" -ArgumentList "$($session.SessionID)","/TIME:1","ScreenConnect settings updated" -WindowStyle Hidden -ErrorAction SilentlyContinue
    }
} catch {
    Write-Verbose "Could not notify all sessions: $($_.Exception.Message)"
}

Write-Host "ScreenConnect tray icon set to '$desiredVal' for all users (current and future)" -ForegroundColor Green
Write-Host "New users will automatically inherit these settings" -ForegroundColor Yellow