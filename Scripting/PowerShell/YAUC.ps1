param (
    [switch]$Yes
)

# Define the KB update to check
$KBToCheck = "KB5048685"

Write-Host "Checking for the applicability of $KBToCheck..." -ForegroundColor Cyan

# Function to check if the update is already installed
function Is-UpdateInstalled {
    param (
        [string]$KBNumber
    )
    $UpdatesInstalled = Get-HotFix | Where-Object { $_.HotFixID -eq $KBNumber }
    if ($UpdatesInstalled) {
        Write-Host "$KBNumber is already installed." -ForegroundColor Green
        return $true
    } else {
        Write-Host "$KBNumber is not installed." -ForegroundColor Yellow
        return $false
    }
}

# Function to install the update
function Install-Update {
    param (
        [object]$Update
    )
    Write-Host "Starting the installation of $KBToCheck..." -ForegroundColor Cyan
    $Installer = New-Object -ComObject Microsoft.Update.Installer
    $Installer.Updates = $Update
    $Result = $Installer.Install()

    if ($Result.ResultCode -eq 2) {
        Write-Host "$KBToCheck installed successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to install $KBToCheck. Result Code: $($Result.ResultCode)" -ForegroundColor Red
    }
}

# Check if the KB update is installed
$IsInstalled = Is-UpdateInstalled -KBNumber $KBToCheck

# If not installed, check if it's available to be installed
if (-not $IsInstalled) {
    # Create the Windows Update Session
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    # Search for the KB update
    Write-Host "Searching for $KBToCheck in available updates..." -ForegroundColor Cyan
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 AND Type='Software'")
    $IsAvailable = $false

    For ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
        $Update = $SearchResult.Updates.Item($i)
        if ($Update.Title -match $KBToCheck) {
            Write-Host "$KBToCheck is available for installation." -ForegroundColor Green
            Write-Host "Details of the update:" -ForegroundColor Cyan
            Write-Host "------------------------------------------------------------"
            Write-Host "Title       : $($Update.Title)"
            Write-Host "Description : $($Update.Description)"
            Write-Host "KB Articles : $($Update.KBArticleIDs -join ', ')"
            Write-Host "------------------------------------------------------------"

            # Prompt user for installation
            if ($Yes -or (Read-Host "Would you like to install $KBToCheck now? (Yes/No)" -eq "Yes")) {
                # Begin installation
                $UpdateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
                $UpdateCollection.Add($Update) | Out-Null
                Install-Update -Update $UpdateCollection
            } else {
                Write-Host "Installation of $KBToCheck was skipped." -ForegroundColor Yellow
            }
            $IsAvailable = $true
            break
        }
    }

    if (-not $IsAvailable) {
        Write-Host "$KBToCheck is not available for installation on this system." -ForegroundColor Red
    }
}

# Additional Check: Validate if a reboot is required
if (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue) {
    Write-Host "A reboot is required to complete a previous update." -ForegroundColor Red
}
