<#
.SYNOPSIS
    Pulls Microsoft Graph logs (Audit, SignIn, or Both) for a specified timeframe,
    excludes SigninLogs.Read.All scope, and gracefully handles the tenant no-premium scenario.

.DESCRIPTION
    - Installs/Imports only Microsoft.Graph.Authentication if needed, letting
      auto-load handle the specific modules for directory audit & sign-in.
    - Allows -LogType "Audit", "SignIn", or "Both"
    - Exports logs to separate CSVs in the script directory
    - If the tenant lacks premium licensing for sign-in logs, the script
      prints a warning and skips sign-in logs.

.PARAMETER LogType
    "Audit", "SignIn", or "Both"

.PARAMETER Days
    Number of days to look back from the current time (default 30).

.NOTES
    Version: 1.4
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet("Audit", "SignIn", "Both")]
    [string]$LogType = "Audit",

    [int]$Days = 30
)

###############################################################################
# 0. Setup & Basic Validation
###############################################################################
Write-Host "=== Incident Response Log Retrieval (No Premium Sign-In Handling) ===" -ForegroundColor Cyan

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "ERROR: This script requires PowerShell 5.0 or higher. Exiting."
    exit 1
}

$StartDate = (Get-Date).AddDays(-$Days)
$EndDate   = Get-Date
Write-Host "Retrieving [$LogType] logs from $($StartDate.ToString()) to $($EndDate.ToString())..."

###############################################################################
# 1. Ensure Microsoft.Graph.Authentication Module
###############################################################################
function Ensure-Module {
    param(
        [Parameter(Mandatory=$true)] [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing $ModuleName because it was not found..."
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
        }
        catch {
            Write-Host "ERROR: Failed to install $ModuleName : $($_.Exception.Message)"
            exit 2
        }
    }
}

Ensure-Module -ModuleName "Microsoft.Graph.Authentication"
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

###############################################################################
# 2. Connect to MgGraph (WITHOUT SigninLogs.Read.All)
###############################################################################
try {
    # We'll rely on "AuditLog.Read.All" + "Directory.Read.All"
    # Typically, sign-in logs also require a premium license, but sometimes
    # tenant roles let you read them with just AuditLog.Read.All.
    $requiredScopes = @("AuditLog.Read.All","Directory.Read.All")

    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
}
catch {
    Write-Host "ERROR: Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 3
}

###############################################################################
# 3. Functions to Retrieve Logs (wrapping single objects in arrays)
###############################################################################
function Get-AuditLogs {
    param(
        [datetime]$Since,
        [datetime]$Until
    )

    $startStr = $Since.ToString("o")
    $endStr   = $Until.ToString("o")

    Write-Host "Pulling Directory Audit logs..."
    try {
        $raw = Get-MgAuditLogDirectoryAudit -All -Filter "activityDateTime ge $startStr and activityDateTime le $endStr"
        return @($raw)  # wrap to handle single-object returns
    }
    catch {
        Write-Warning "Failed to retrieve Audit logs: $($_.Exception.Message)"
        return $null
    }
}

function Get-SignInLogs {
    param(
        [datetime]$Since,
        [datetime]$Until
    )

    $startStr = $Since.ToString("o")
    $endStr   = $Until.ToString("o")

    Write-Host "Pulling Sign-In logs..."
    try {
        $raw = Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $startStr and createdDateTime le $endStr"
        return @($raw)  # wrap to handle single-object returns
    }
    catch {
        # Check if it's the premium license error:
        if ($_.Exception.Message -match "Tenant is not a B2C tenant and doesn't have premium license") {
            Write-Warning "Sign-In logs are unavailable: tenant lacks premium license. Skipping sign-in logs."
        }
        else {
            Write-Warning "Failed to retrieve Sign-In logs: $($_.Exception.Message)"
        }
        return $null
    }
}

###############################################################################
# 4. Export Logs Function
###############################################################################
function Export-Logs {
    param(
        [Parameter(Mandatory=$true)] [string]$LogCategory,  # "Audit" or "SignIn"
        [Parameter(Mandatory=$true)] $Data
    )

    if (-not $Data) {
        Write-Host "No $LogCategory data to export."
        return
    }

    $count = $Data.Count
    Write-Host "Exporting $count $LogCategory record(s)..."

    $timestamp = (Get-Date -Format "yyyyMMdd_HHmmss")
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    if (-not $scriptDir) { $scriptDir = $PWD }

    $fileName  = "$LogCategory-Logs_$timestamp.csv"
    $outPath   = Join-Path $scriptDir $fileName

    try {
        $Data | Export-Csv -Path $outPath -NoTypeInformation
        Write-Host "Saved $LogCategory logs to: $outPath"
    }
    catch {
        Write-Host "ERROR: Failed to export $LogCategory logs: $($_.Exception.Message)"
    }
}

###############################################################################
# 5. Retrieve & Export
###############################################################################
Write-Host "`n==== Retrieving data (this may take a while) ===="
switch ($LogType) {
    "Audit" {
        $auditData = Get-AuditLogs -Since $StartDate -Until $EndDate
        Export-Logs -LogCategory "Audit" -Data $auditData
    }
    "SignIn" {
        $signInData = Get-SignInLogs -Since $StartDate -Until $EndDate
        Export-Logs -LogCategory "SignIn" -Data $signInData
    }
    "Both" {
        $auditData  = Get-AuditLogs   -Since $StartDate -Until $EndDate
        $signInData = Get-SignInLogs -Since $StartDate -Until $EndDate

        Export-Logs -LogCategory "Audit"  -Data $auditData
        Export-Logs -LogCategory "SignIn" -Data $signInData
    }
}

###############################################################################
# 6. Disconnect & Wrap Up
###############################################################################
Disconnect-MgGraph | Out-Null
Write-Host "`nDisconnected from Microsoft Graph."
Write-Host "=== Script complete. ==="
exit 0
