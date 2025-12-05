<#
.SYNOPSIS
    Retrieves all Azure AD users where OnPremisesSyncEnabled = True,
    correlates them to on-prem AD (first by sAMAccountName match, then by DisplayName guess),
    then exports a CSV with:
      AAD_UPN, AAD_DisplayName, OnPrem_UPN, OnPrem_FirstName, OnPrem_LastName, OnPrem_DistinguishedName

.DESCRIPTION
    - Connects to Microsoft Graph, gets all users with OnPremisesSyncEnabled = $true.
    - For each, parse local part of UPN => sAMAccountName. 
      If that fails, split DisplayName into "first + last" tokens, attempt (givenName, sn) match.
    - The script then retrieves userprincipalname, givenname, sn, and especially "distinguishedName" from AD,
      to show the full path "CN=...,OU=...,DC=..." for that user.

.NOTES
    - Must run with domain privileges to read AD, and Graph access to read all user objects.
    - If multiple or zero AD matches appear, we skip or warn. 
#>

###############################################################################
# 1) Connect to Microsoft Graph
###############################################################################
Connect-MgGraph -Scopes "User.Read.All"

Write-Host "Retrieving Azure AD users who have OnPremisesSyncEnabled = True..."

# 2) Retrieve all AAD users with that flag
$allAzureUsers = Get-MgUser -All -Property "DisplayName","UserPrincipalName","OnPremisesSyncEnabled" `
    | Where-Object { $_.OnPremisesSyncEnabled -eq $true }

Write-Host "Found $($allAzureUsers.Count) such AAD users."

###############################################################################
# 3) We'll store final correlation data here
###############################################################################
$results = @()

###############################################################################
# 4) Helper function: search by sAMAccountName
###############################################################################
function Try-SamAccountName($sam) {
    $searcher = [ADSISearcher]"(sAMAccountName=$($sam))"
    $searcher.PageSize = 5000
    return $searcher.FindOne()
}

###############################################################################
# 5) Helper function: search by (givenName=xxx)(sn=yyy)
###############################################################################
function Try-GivenNameSn($given, $sn) {
    $filter = "(&(objectClass=user)(givenName=$($given))(sn=$($sn)))"
    $searcher = [ADSISearcher]$filter
    $searcher.PageSize = 5000
    return $searcher.FindAll()
}

###############################################################################
# 6) Main loop over each AAD user
###############################################################################
foreach ($user in $allAzureUsers) {

    $azureUPN         = $user.UserPrincipalName
    $azureDisplayName = $user.DisplayName

    # Extract local part of UPN => potential sAMAccountName
    $localPart = $null
    if ($azureUPN -match '@') {
        $localPart = $azureUPN.Split('@')[0]
    }
    else {
        # fallback if there's no domain portion
        $localPart = $azureUPN
    }

    # Prepare placeholders
    $onPremUPN       = $null
    $onPremFirstName = $null
    $onPremLastName  = $null
    $onPremDN        = $null  # We'll store distinguishedName here

    ###########################################################################
    # 6a) Attempt sAMAccountName = localPart
    ###########################################################################
    $found = $null
    if ($localPart) {
        $found = Try-SamAccountName $localPart
    }

    if ($found) {
        # DistName
        $onPremDN        = $found.Properties['distinguishedname']
        $onPremUPN       = $found.Properties['userprincipalname']
        $onPremFirstName = $found.Properties['givenname']
        $onPremLastName  = $found.Properties['sn']
    }
    else {
        # 6b) Attempt fallback parse of DisplayName => (first, last)
        if ($azureDisplayName -and $azureDisplayName.Contains(" ")) {
            $tokens = $azureDisplayName.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
            if ($tokens.Count -ge 2) {
                $guessedFirst = $tokens[0]
                $guessedLast  = $tokens[$tokens.Count - 1]

                $matches = Try-GivenNameSn $guessedFirst $guessedLast

                if ($matches -and $matches.Count -eq 1) {
                    # We'll pick this single match
                    $item = $matches[0]
                    $onPremDN        = $item.Properties['distinguishedname']
                    $onPremUPN       = $item.Properties['userprincipalname']
                    $onPremFirstName = $item.Properties['givenname']
                    $onPremLastName  = $item.Properties['sn']
                }
                elseif ($matches.Count -gt 1) {
                    Write-Warning "Multiple on-prem AD matches for name '$guessedFirst $guessedLast' => skipping user $azureUPN"
                }
                else {
                    Write-Warning "No on-prem AD user found for guessed name '$guessedFirst $guessedLast' => $azureUPN"
                }
            }
            else {
                Write-Warning "DisplayName '$azureDisplayName' has fewer than 2 tokens => skipping name fallback for $azureUPN"
            }
        }
        else {
            Write-Warning "No space in DisplayName or no DisplayName => skipping name fallback for $azureUPN"
        }
    }

    # Convert arrays to strings (distinguishedName is often an array of 1)
    if ($onPremDN -is [System.Collections.IEnumerable]) {
        $onPremDN = $onPremDN -join '; '
    }
    if ($onPremFirstName -is [System.Collections.IEnumerable]) {
        $onPremFirstName = $onPremFirstName -join '; '
    }
    if ($onPremLastName -is [System.Collections.IEnumerable]) {
        $onPremLastName = $onPremLastName -join '; '
    }
    if ($onPremUPN -is [System.Collections.IEnumerable]) {
        $onPremUPN = $onPremUPN -join '; '
    }

    ###########################################################################
    # 6c) Add result row
    ###########################################################################
    $results += [pscustomobject]@{
        AAD_UPN              = $azureUPN
        AAD_DisplayName      = $azureDisplayName
        OnPrem_UPN           = $onPremUPN
        OnPrem_FirstName     = $onPremFirstName
        OnPrem_LastName      = $onPremLastName
        OnPrem_DistinguishedName = $onPremDN
    }
}

###############################################################################
# 7) Export to CSV
###############################################################################
$csvPath = "C:\Techneaux\AAD_OnPrem_Correlation_DistinguishedName.csv"
$results | Export-Csv -NoTypeInformation -Path $csvPath

Write-Host "`nExported $($results.Count) records to $csvPath. Done."