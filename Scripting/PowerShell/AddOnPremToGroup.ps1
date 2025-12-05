<#
.SYNOPSIS
    Add a list of on-prem AD users (by their distinguishedName) to a given security group.

.DESCRIPTION
    - Reads a CSV that has a column "OnPrem_DistinguishedName".
    - Locates the group "ProvisionedToHybrid" in on-prem AD.
    - For each user DN, attempts groupEntry.Add("LDAP://<userDN>").

.NOTES
    - Must run with sufficient permissions to modify group membership in AD.
#>

# 1) Path to your CSV from the prior correlation script:
$csvPath = "C:\Techneaux\AAD_OnPrem_Correlation_DistinguishedName.csv"

# 2) Column name that contains the user DN:
$userDNColumn = "OnPrem_DistinguishedName"

# 3) Group name to add them to:
$groupName = "ProvisionedToHybrid"

# 4) Load the data
Write-Host "Reading user list from CSV: $csvPath..."
$users = Import-Csv -Path $csvPath

Write-Host "Searching AD for group '$groupName'..."
$searcher = [ADSISearcher]"(&(objectClass=group)(cn=$($groupName)))"
$found = $searcher.FindOne()

if (-not $found) {
    Write-Warning "Could not find group named '$groupName' in AD. Exiting."
    return
}

$groupDN = $found.Properties['distinguishedname']
if ($groupDN -is [System.Collections.IEnumerable]) {
    $groupDN = $groupDN -join '; '
}
Write-Host "Found group '$groupName' with DN: $groupDN"

# 5) Bind to the group as a DirectoryEntry
$groupAdsPath = "LDAP://$groupDN"
$groupEntry = [ADSI]$groupAdsPath

# 6) Loop over each user row in the CSV
foreach ($row in $users) {
    $userDN = $row.$userDNColumn
    if (-not $userDN) {
        Write-Warning "No user DN found in column '$userDNColumn' for row $($row.AAD_UPN). Skipping..."
        continue
    }

    # Attempt to add them
    $userAdsPath = "LDAP://$userDN"
    try {
        Write-Host "Adding $userDN to group $groupDN..."
        $groupEntry.psbase.Invoke("Add", $userAdsPath)
    }
    catch {
        Write-Warning "Failed to add $userDN to group $groupDN : $($_.Exception.Message)"
    }
}

Write-Host "Done adding users to $groupName."