# Check if the Active Directory module is available and import it
try {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host "The Active Directory module is not installed. Please install it using RSAT or relevant tools." -ForegroundColor Red
        exit
    }
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host "Active Directory module imported successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to import the Active Directory module. Please ensure it is installed and try again." -ForegroundColor Red
    exit
}

# Define required policy values
$requiredPolicy = @{
    MinPasswordLength   = 16
    ComplexityRequired  = $true
    MaxPasswordAge      = 365 # in days
    LockoutThreshold    = 5   # Example threshold for failed attempts
}

# Function to evaluate password length and complexity
function Check-PasswordPolicy {
    $passwordPolicy = Get-ADDefaultDomainPasswordPolicy

    $result = @(
        @{
            Setting     = "Minimum Password Length"
            Expected    = "$($requiredPolicy.MinPasswordLength)+ Characters"
            Found       = "$($passwordPolicy.MinPasswordLength) Characters"
            Compliant   = $passwordPolicy.MinPasswordLength -ge $requiredPolicy.MinPasswordLength
        },
        @{
            Setting     = "Password Complexity Enabled"
            Expected    = "3+ Mixed Characters"
            Found       = $(if ($passwordPolicy.ComplexityEnabled) { "Enabled" } else { "Disabled" })
            Compliant   = $passwordPolicy.ComplexityEnabled -eq $requiredPolicy.ComplexityRequired
        },
        @{
            Setting     = "Maximum Password Age"
            Expected    = "$($requiredPolicy.MaxPasswordAge) Days"
            Found       = "$($passwordPolicy.MaxPasswordAge.TotalDays) Days"
            Compliant   = $passwordPolicy.MaxPasswordAge.TotalDays -le $requiredPolicy.MaxPasswordAge
        }
    )

    return $result
}

# Function to evaluate lockout policy
function Check-LockoutPolicy {
    $passwordPolicy = Get-ADDefaultDomainPasswordPolicy

    $result = @(
        @{
            Setting     = "Account Lockout Threshold"
            Expected    = "$($requiredPolicy.LockoutThreshold)+ failed attempts"
            Found       = $(if ($passwordPolicy.LockoutThreshold) { "$($passwordPolicy.LockoutThreshold) failed attempts" } else { "Not Configured" })
            Compliant   = $passwordPolicy.LockoutThreshold -ge $requiredPolicy.LockoutThreshold
        }
    )

    return $result
}

# Collect results
$passwordResults = Check-PasswordPolicy
$lockoutResults = Check-LockoutPolicy

# Combine all results
$allResults = $passwordResults + $lockoutResults

# Generate file name dynamically based on hostname and current date
$hostname = $env:COMPUTERNAME
$date = (Get-Date -Format "yyyy-MM-dd")
$outputFile = "C:\${hostname}__PasswordCompliance--${date}.txt"

# Generate report content
$outputContent = @"
=========================================
GPO Policy Compliance Report
Hostname: $hostname
Generated on: $(Get-Date)
=========================================

Checklist Table (Expected Values):
-----------------------------------
$($allResults | ForEach-Object {
    "Setting: $($_.Setting)`nExpected: $($_.Expected)`n"
})

Values Found Table:
-------------------
$($allResults | ForEach-Object {
    "Setting: $($_.Setting)`nFound: $($_.Found)`n"
})

Compliance Table:
-----------------
$($allResults | ForEach-Object {
    "Setting: $($_.Setting)`nCompliant: $(if ($_.Compliant) { 'Yes' } else { 'No' })`n"
})
"@

# Write report to file
$outputContent | Out-File -FilePath $outputFile -Encoding UTF8

# Print report location
Write-Host "`nThe GPO Policy Compliance Report has been saved to: $outputFile" -ForegroundColor Cyan

# Optional: Open the report automatically
Start-Process notepad.exe $outputFile
