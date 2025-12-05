<#
    .SYNOPSIS
        Performs a hard match for a single on-prem AD user to Azure AD by setting the ImmutableID.

    .DESCRIPTION
        - Retrieves a specified on-premises AD user (by UPN).
        - Splits the user’s UPN at '@' and re-joins it with @vpsb.net.
        - Retrieves the on-premises ObjectGUID, converts it to Base64 as the ImmutableID.
        - Sets the ImmutableID in Azure AD to achieve a hard match.
        - Outputs mapping (OnPremUser => AzureADUser).

    .NOTES
        - Requires: ActiveDirectory PowerShell module (for Get-ADUser)
        - Requires: MSOnline PowerShell module (for Get-MsolUser, Set-MsolUser)
        - Ensure you have done: Connect-MsolService
        - Use caution and test thoroughly in a non-production environment.

#>

# Customize this variable to test a specific on-prem user:
$TestOnPremUserUPN = 'techneaux@vpss.local'

# Import required modules if needed (comment out if they're already imported elsewhere).
Import-Module ActiveDirectory
Import-Module MSOnline

# Ensure you have already run Connect-MsolService for Azure AD
# Connect-MsolService

Write-Host "Attempting to retrieve On-Prem AD user: $TestOnPremUserUPN"
try {
    # Get the single on-prem user with all necessary properties
    $onPremUser = Get-ADUser -Identity $TestOnPremUserUPN -Properties UserPrincipalName, ObjectGUID -ErrorAction Stop

    # Check that we actually got a user object
    if (-not $onPremUser.UserPrincipalName) {
        Write-Warning "User does not have a valid UPN. Aborting."
        return
    }
    
    # Split the local part of the UPN (everything before '@') and create new UPN with @vpsb.net
    $localPart = $onPremUser.UserPrincipalName.Split('@')[0]
    $azureADUPN = "$localPart@vpsb.net"

    # Convert on-prem AD ObjectGUID to a Base64 string
    $immutableID = [System.Convert]::ToBase64String($onPremUser.ObjectGUID.ToByteArray())

    Write-Host "Computed ImmutableID (Base64 of ObjectGUID): $immutableID"

    # Attempt to locate this user in Azure AD and set the ImmutableID
    $azureUser = Get-MsolUser -UserPrincipalName $azureADUPN -ErrorAction Stop

    Write-Host "Found Azure AD user $azureADUPN. Setting ImmutableID..."
    Set-MsolUser -UserPrincipalName $azureADUPN -ImmutableId $immutableID -ErrorAction Stop

    Write-Host "Successfully set ImmutableID:"
    Write-Host "$($onPremUser.UserPrincipalName) => $azureADUPN"
}
catch {
    Write-Warning "Failed to process user '$TestOnPremUserUPN': $($_.Exception.Message)"
}
