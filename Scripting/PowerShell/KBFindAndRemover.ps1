<#
.SYNOPSIS
    Script to remove a specific KB update using PSWindowsUpdate.
    Installs PSWindowsUpdate if needed, removes specified KB, then uninstalls PSWindowsUpdate if it was originally absent.

.DESCRIPTION
    1. Sets execution policy to RemoteSigned.
    2. Checks if PSWindowsUpdate module is installed.
    3. If not installed, installs PSWindowsUpdate, imports it, and flags that we did so.
    4. Checks for the presence of the specified KB via Get-HotFix.
    5. If present, removes it using Remove-WindowsUpdate; otherwise, reports that it is not installed.
    6. If the script itself installed PSWindowsUpdate, it uninstalls the module upon completion.
    7. Logs basic status messages.

.PARAMETER KBID
    Required parameter; the KB article ID (e.g., "KB5005565") to remove.

.NOTES
    Please run this script with elevated privileges (Run as Administrator).
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^KB\\d+$",
        ErrorMessage = "The KBID must be in the format 'KB' followed by digits (e.g., KB1234567).")]
    [string]$KBID
)

Write-Host "=== Starting KB Removal Script ==="
Write-Host "You have specified KB: $KBID"

try {
    # Step 1: Set execution policy (this affects the local machine scope)
    # You may choose -Scope Process to avoid permanently changing the system-wide policy.
    $initExecPolicy = Get-ExecutionPolicy
    Write-Host "Checking and potentially changing Execution Policy..."
    if ( $initExecPolicy -ne "RemoteSigned" ) 
    {
        Set-ExecutionPolicy RemoteSigned -Force
        $execPolChange = 1
    }

    # Step 2: Check if PSWindowsUpdate is already installed
    Write-Host "Checking if PSWindowsUpdate is already installed..."
    $ModuleName       = "PSWindowsUpdate"
    $moduleInstalled  = Get-Module -Name $ModuleName -ListAvailable
    $installedByScript = $false

    if (-not $moduleInstalled) {
        Write-Host "PSWindowsUpdate is NOT installed. Installing..."
        Install-Module -Name $ModuleName -Force

        # Mark that we installed this module ourselves, so we can safely remove later.
        $installedByScript = $true
    } else {
        Write-Host "PSWindowsUpdate is already installed on this system."
    }

    # Step 3: Import the module
    Write-Host "Importing PSWindowsUpdate module..."
    Import-Module PSWindowsUpdate -ErrorAction Stop

    # Optional - Show version, available commands, etc. (uncomment as desired)
    # Write-Host "Module version information:"
    # Get-Package -Name $ModuleName
    # Write-Host "Available PSWindowsUpdate Commands:"
    # Get-Command -Module PSWindowsUpdate

    # Step 4: Check if the KB is installed
    Write-Host "Checking if $KBID is installed on this system..."
    $kbCheck = Get-HotFix -Id $KBID -ErrorAction SilentlyContinue

    if ($kbCheck) {
        Write-Host "$KBID is installed. Proceeding with removal..."

        # Step 5: Remove the KB using PSWindowsUpdate
        #  NOTE: The -NoRestart switch prevents automatic restarts
        #        If you want to allow automatic restarts, remove -NoRestart.
        Remove-WindowsUpdate -KBArticleID $KBID -NoRestart -Confirm:$false

        Write-Host "$KBID removal complete (pending any necessary restarts)."
    } else {
        Write-Host "$KBID is NOT installed on this system."
    }

    # Step 6: If we installed PSWindowsUpdate ourselves, uninstall it
    if ($installedByScript) {
        Write-Host "Uninstalling PSWindowsUpdate module (since we installed it)..."
        # Remove-Module to unload it from the current session
        Remove-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue

        # Uninstall-Module to remove from the system
        Uninstall-Module -Name $ModuleName -AllVersions -Force -ErrorAction SilentlyContinue

        Write-Host "PSWindowsUpdate module uninstalled."
    }
    if ($execPolChange)
    {
        Set-ExecutionPolicy $initExecPolicy -Force
    }
    Write-Host "=== Script Execution Complete ==="
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
    Write-Host "Script is terminating."
    exit 1
}