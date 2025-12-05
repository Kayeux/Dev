<#
.SYNOPSIS
    Performs a hard match for each user in On-Prem AD to Azure AD by setting the OnPremisesImmutableId using Microsoft Graph.
    Additionally:
     - Sets extensionAttribute13 = "ScriptSynced" on success to avoid re-processing those users (unless -ReSync is set).
     - Has a global script timer that prints total runtime at completion or on error exit.
     - Shows a progress metric (x / y) for normal runs.
     - Provides two override params (-OnPremUPN and -AADUPN) to do a one-time direct sync for a single user, skipping all normal logic.

.DESCRIPTION
    - Retrieves on-premises AD users (Get-ADUser), skipping those in Azure AD that are already flagged as "ScriptSynced" (unless -ReSync).
    - Constructs an Azure AD UPN from the on-prem UPN and the provided PrimaryDomain.
    - Converts the on-prem AD ObjectGUID to a Base64 string (ImmutableID).
    - Retrieves all Azure AD users once and builds a dictionary for fast lookups.
    - If the standard matching (primary guess + fallback patterns) fails,
      and if -UseNicknames is specified, attempts to map the on-prem FIRST NAME to synonyms from the provided (or default) file.
      (But that logic is skipped if you do the single-user override using -OnPremUPN/-AADUPN.)
    - Logs successes and failures to CSV, prepending domain discrepancy once if needed.
    - Produces a trimmed orphan report with (GivenName, OnPremisesImmutableId, Surname, UPN).
    - If both -OnPremUPN and -AADUPN are set, we skip normal logic and try to update that single user only.

.NOTES
    - Requires: ActiveDirectory and Microsoft Graph modules.
    - Run with an account that can install modules and authenticate to Graph.
    - Test thoroughly in a non-production environment.
#>

param(
    [string]$FailureCsvPath       = (Join-Path $PSScriptRoot "Logs\DirSyncUPNHardMatch_Failure_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"),
    [string]$SuccessCsvPath       = (Join-Path $PSScriptRoot "Logs\DirSyncUPNHardMatch_Success_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"),
    [string]$OrphanReportCsvPath  = (Join-Path $PSScriptRoot "Logs\AzureNotInAD_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"),

    [int]$MaxUsers                = 20,
    [bool]$FullSend               = $false,

    [string]$PrimaryDomain        = $null,
    [string]$MSOLDomain           = $null,

    # If false, we skip users who have extensionAttribute13 = "ScriptSynced"
    # If true, we ignore that attribute and re-process them
    [bool]$ReSync                 = $false,

    # By default, -UseNicknames is OFF
    [bool]$UseNicknames           = $false,

    # By default, NicknameFile is "nicknames.txt" in the same folder as this script
    [string]$NicknameFile = (Join-Path (Split-Path $MyInvocation.MyCommand.Path) "nicknames.txt"),

    # If BOTH of these are provided, skip normal logic and do a single direct sync
    [string]$OnPremUPN            = $null,
    [string]$AADUPN               = $null
)

# Ensure log directory exists
$logDirectory = "$PSScriptRoot\Logs"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

###############################################################################
# 0. Script Timer & Helper for Early Exit
###############################################################################
$scriptStart = Get-Date
function StopAndExit($code) {
    $elapsed = (Get-Date) - $scriptStart
    Write-Host "Script ended. Total runtime: $elapsed"
    exit $code
}

###############################################################################
# 1. Validate required parameters (normal mode)
###############################################################################
# If the user provided only OnPremUPN or only AADUPN, fail
if (($OnPremUPN -and -not $AADUPN) -or ($AADUPN -and -not $OnPremUPN)) {
    Write-Warning "Both -OnPremUPN and -AADUPN must be provided together, or neither. Exiting..."
    StopAndExit 7
}

# If both OnPremUPN and AADUPN are set, we skip all normal logic below (after basic module loads & connect)
$singleUserOverride = $false
if ($OnPremUPN -and $AADUPN) {
    $singleUserOverride = $true
}

# If NOT singleUserOverride, proceed with normal param checks
if (-not $singleUserOverride) {
    if (-not $PrimaryDomain) {
        Write-Host "ERROR: You MUST include the Org's Primary Domain. Example: -PrimaryDomain techneaux.com"
        StopAndExit 1
    }
}

###############################################################################
# 2. Install Microsoft.Graph module if needed
###############################################################################
try {
    Install-Module Microsoft.Graph -ErrorAction Ignore -Scope AllUsers
}
catch {
    Write-Host "ERROR: Microsoft Graph Module Failed to Install. Troubleshoot manually then re-run script."
    StopAndExit 2
}

###############################################################################
# 3. Import required modules
###############################################################################
Write-Host "Importing required modules (if not already loaded)..."
Import-Module ActiveDirectory
Import-Module Microsoft.Graph.Users

###############################################################################
# 4. Connect to Microsoft Graph
###############################################################################
try {
    Write-Host "Attempting to Connect to MgGraph. Be ready to authenticate with GA credentials..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All"
}
catch {
    Write-Host "ERROR: Failed to Connect to MgGraph. Troubleshoot manually then re-run."
    StopAndExit 3
}

###############################################################################
# If single-user override is set, do that logic here and then exit
###############################################################################
if ($singleUserOverride) {
    Write-Host "Single-user override mode: linking On-Prem UPN '$OnPremUPN' to Azure UPN '$AADUPN'..."

    # 1) Retrieve the on-prem user to get their ObjectGUID
    try {
        $onPremUserObj = Get-ADUser -Filter { UserPrincipalName -eq $OnPremUPN } -Properties ObjectGUID, Enabled
    }
    catch {
        Write-Warning "Failed to retrieve OnPrem user $OnPremUPN : $($_.Exception.Message)"
        StopAndExit 11
    }

    if (-not $onPremUserObj) {
        Write-Warning "OnPrem user $OnPremUPN not found. Exiting..."
        StopAndExit 12
    }

    # 2) Convert their GUID to Base64
    $immutableId = [System.Convert]::ToBase64String($onPremUserObj.ObjectGUID.ToByteArray())

    # 3) Attempt to update that AAD user
    #    We won't retrieve the entire user object from Azure; we'll just call an update
    #    If they don't exist or can't be updated, that'll fail with an exception
    try {
        # Set the OnPremisesImmutableId
        Update-MgUser -UserId $AADUPN -OnPremisesImmutableId $immutableId -ErrorAction Stop

        # Mark extensionAttribute13 = "ScriptSynced"
        Update-MgUser -UserId $AADUPN -OnPremisesExtensionAttributes @{ ExtensionAttribute13 = "ScriptSynced" } -ErrorAction Stop

        Write-Host "SUCCESS: Single-user override: Set OnPremisesImmutableId ($immutableId) for $OnPremUPN => $AADUPN"
    }
    catch {
        Write-Warning "Failed to update $AADUPN : $($_.Exception.Message)"
        StopAndExit 13
    }

    Write-Host "Single-user linking completed successfully."

    # Print final timer & exit
    $elapsed = (Get-Date) - $scriptStart
    Write-Host "Script completed successfully in $elapsed"
    Disconnect-MgGraph | Out-Null
    exit 0
}

###############################################################################
# 5. Verify PrimaryDomain against tenant’s default domain (normal mode only)
###############################################################################
try {
    $org = Get-MgOrganization -ErrorAction Stop
    $defaultDomain = ($org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
}
catch {
    Write-Host "ERROR: Could not retrieve organization domain information."
    StopAndExit 5
}

$domainDiscrepancy = ""
if ($PrimaryDomain.ToLower() -ne $defaultDomain.ToLower()) {
    Write-Host "WARNING: The PrimaryDomain you provided ('$PrimaryDomain') does not match the tenant's default domain ('$defaultDomain')."
    $response = Read-Host "Would you like to continue? (Y/N)"
    if ($response -notin @("Y","y")) {
        Write-Host "User opted to exit due to domain mismatch."
        StopAndExit 6
    }
    else {
        Write-Host "Continuing execution. Discrepancy logged."
        $domainDiscrepancy = "User provided PrimaryDomain ('$PrimaryDomain') does not match tenant default domain ('$defaultDomain')."
    }
}
else {
    Write-Host "PrimaryDomain matches the tenant default domain: $PrimaryDomain"
}

###############################################################################
# 6. Load Nickname File if -UseNicknames
###############################################################################
$NameSynonyms = @{}  # dictionary for first-name synonyms

if ($UseNicknames) {
    if (Test-Path $NicknameFile) {
        Write-Host "Using nicknames from file: $NicknameFile"

        $lines = Get-Content -Path $NicknameFile
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }  # skip empty lines
            $rawNames = $line -split '\s+' | Where-Object { $_ -ne "" }
            if (-not $rawNames) { continue }
            foreach ($oneName in $rawNames) {
                $lowerName = $oneName.ToLower()
                if (-not $NameSynonyms.ContainsKey($lowerName)) {
                    $NameSynonyms[$lowerName] = New-Object System.Collections.Generic.List[string]
                }
                foreach ($other in $rawNames) {
                    if ($other -ne $oneName) {
                        $NameSynonyms[$lowerName].Add($other)
                    }
                }
            }
        }
    }
    else {
        Write-Host "WARNING: -UseNicknames was specified but '$NicknameFile' does not exist. Nickname logic will be skipped."
    }
}

###############################################################################
# 7. Retrieve On-Prem AD users
###############################################################################
if ($FullSend) {
    Write-Host "FULL SEND: Retrieving ALL On-Prem AD users..."
    $OnPremUsers = Get-ADUser -Filter * -Properties UserPrincipalName, ObjectGUID, GivenName, Surname, MiddleName, Initials, Enabled
}
else {
    Write-Host "Retrieving up to $MaxUsers On-Prem AD users..."
    $OnPremUsers = Get-ADUser -Filter * -Properties UserPrincipalName, ObjectGUID, GivenName, Surname, MiddleName, Initials, Enabled |
        Select-Object -First $MaxUsers
}

###############################################################################
# 8. Retrieve all Azure AD users once and build a dictionary for fast lookup
###############################################################################
Write-Host "Retrieving all Azure AD users..."
try {
    # We also need OnPremisesExtensionAttributes to see extensionAttribute13
    $AllAzureUsers = Get-MgUser -All -Property "GivenName","Surname","OnPremisesImmutableId",
                                     "UserPrincipalName","OnPremisesExtensionAttributes"
}
catch {
    Write-Host "ERROR: Failed to retrieve Azure AD users. Exiting."
    StopAndExit 4
}

# If not ReSync, skip those who are already ScriptSynced
if (-not $ReSync) {
    Write-Host "Filtering out Azure AD users who are already 'ScriptSynced' in extensionAttribute13..."
    $AllAzureUsers = $AllAzureUsers | Where-Object {
        $_.OnPremisesExtensionAttributes.ExtensionAttribute13 -ne "ScriptSynced"
    }
}

$AzureUsersByUPN = @{}
foreach ($u in $AllAzureUsers) {
    if ($u.UserPrincipalName) {
        $AzureUsersByUPN[$u.UserPrincipalName.ToLower()] = $u
    }
}

###############################################################################
# 9. Prepare a collection to store output results and track processed UPNs
###############################################################################
$Results       = @()
$ProcessedUPNs = @()

###############################################################################
# 10. Process each On-Prem AD user (with progress metric)
###############################################################################
$totalUsers = $OnPremUsers.Count
$currentUserIndex = 0

foreach ($user in $OnPremUsers) {
    $currentUserIndex++

    # Progress preamble
    Write-Host "[ $currentUserIndex / $totalUsers ] Attempting user $($user.UserPrincipalName)..."

    # Skip if no UPN exists
    if (-not $user.UserPrincipalName) {
        Write-Warning "[ $currentUserIndex / $totalUsers ] Skipping user $($user.SamAccountName) - no UPN found."
        continue
    }

    # Gather on-prem name details
    $onPremFirstName = $user.GivenName
    $onPremLastName  = $user.Surname
    $onPremMiddle    = $user.MiddleName

    # Optional parse SamAccountName if needed
    if ((-not $onPremFirstName -or -not $onPremLastName) -and $user.SamAccountName -and $user.SamAccountName.Contains(" ")) {
        $parts = $user.SamAccountName.Split(" ", 2)
        if ($parts.Count -eq 2) {
            $onPremFirstName = $parts[0]
            $onPremLastName  = $parts[1]
            Write-Verbose "Parsed SamAccountName '$($user.SamAccountName)' => '$onPremFirstName' + '$onPremLastName'"
        }
    }

    # Build a primary guess for Azure AD UPN
    $localPart       = $user.UserPrincipalName.Split('@')[0]
    $primaryGuessUPN = "$localPart@$PrimaryDomain"

    $azureUser       = $AzureUsersByUPN[$primaryGuessUPN.ToLower()]
    $finalMatchedUPN = $null

    # If the primary guess is found, great
    if ($azureUser) {
        $finalMatchedUPN = $primaryGuessUPN
    }
    else {
        # Fallback patterns if we have some first/last
        if ($onPremFirstName -and $onPremLastName) {
            $candidateUPNs = @()
            $candidateUPNs += "$onPremFirstName@$PrimaryDomain"
            if ($onPremFirstName.Length -ge 1 -and $onPremLastName.Length -ge 1) {
                $candidateUPNs += ("{0}{1}@{2}" -f $onPremFirstName.Substring(0,1), $onPremLastName, $PrimaryDomain)
            }
            if ($onPremMiddle -and $onPremMiddle.Length -ge 1) {
                $candidateUPNs += ("{0}{1}{2}@{3}" -f $onPremFirstName.Substring(0,1), $onPremMiddle.Substring(0,1), $onPremLastName, $PrimaryDomain)
            }
            if ($onPremLastName.Length -ge 1) {
                $candidateUPNs += ("{0}{1}@{2}" -f $onPremFirstName, $onPremLastName.Substring(0,1), $PrimaryDomain)
            }
            $candidateUPNs += ("{0}.{1}@{2}" -f $onPremFirstName, $onPremLastName, $PrimaryDomain)

            foreach ($candidate in $candidateUPNs | Where-Object { $_ }) {
                Write-Verbose "Trying fallback UPN: $candidate"
                $potentialUser = $AzureUsersByUPN[$candidate.ToLower()]
                if ($potentialUser) {
                    # Check name match
                    if ($potentialUser.GivenName -eq $onPremFirstName -and $potentialUser.Surname -eq $onPremLastName) {
                        $azureUser       = $potentialUser
                        $finalMatchedUPN = $candidate
                        break
                    }
                }
            }
        }

        # Nickname fallback
        if (-not $finalMatchedUPN -and $UseNicknames `
            -and (-not [string]::IsNullOrWhiteSpace($onPremFirstName)) `
            -and (-not [string]::IsNullOrWhiteSpace($onPremLastName))) {

            $firstNameLower = $onPremFirstName.ToLower()
            if ($NameSynonyms.ContainsKey($firstNameLower)) {
                $synonyms = $NameSynonyms[$firstNameLower]
                foreach ($synonym in $synonyms) {
                    $candidateUPNs = @()
                    $candidateUPNs += "$synonym@$PrimaryDomain"
                    if ($onPremLastName.Length -ge 1) {
                        if ($synonym.Length -ge 1) {
                            $candidateUPNs += ("{0}{1}@{2}" -f $synonym.Substring(0,1), $onPremLastName, $PrimaryDomain)
                        }
                        $candidateUPNs += ("{0}{1}@{2}" -f $synonym, $onPremLastName.Substring(0,1), $PrimaryDomain)
                        if ($onPremMiddle -and $onPremMiddle.Length -ge 1 -and $synonym.Length -ge 1) {
                            $candidateUPNs += ("{0}{1}{2}@{3}" -f $synonym.Substring(0,1), $onPremMiddle.Substring(0,1), $onPremLastName, $PrimaryDomain)
                        }
                        $candidateUPNs += ("{0}.{1}@{2}" -f $synonym, $onPremLastName, $PrimaryDomain)
                    }

                    foreach ($candidate in $candidateUPNs | Where-Object { $_ }) {
                        Write-Verbose "Trying nickname-based UPN: $candidate"
                        $potentialUser = $AzureUsersByUPN[$candidate.ToLower()]
                        if ($potentialUser) {
                            if ($potentialUser.Surname -eq $onPremLastName) {
                                Write-Verbose "Matched via nickname fallback: $candidate"
                                $azureUser       = $potentialUser
                                $finalMatchedUPN = $candidate
                                break
                            }
                        }
                    }
                    if ($finalMatchedUPN) { break }
                }
            }
        }
    }

    # If still no match, record failure
    if (-not $azureUser) {
        $Results += [pscustomobject]@{
            OnPremUser         = $user.UserPrincipalName
            AzureADUser        = $primaryGuessUPN
            Status             = "FAILED"
            Reason             = "User not found in Azure AD"
            OnPremEnabled      = $user.Enabled
            NewID              = [System.Convert]::ToBase64String($user.ObjectGUID.ToByteArray())
            PreviousID         = $null
            DomainDiscrepancy  = $domainDiscrepancy
        }
        Write-Warning "[ $currentUserIndex / $totalUsers ] FAILED: $($user.UserPrincipalName) => $primaryGuessUPN : No match in Azure AD"
        continue
    }

    # We have a match: add the final UPN to $ProcessedUPNs
    $ProcessedUPNs += $finalMatchedUPN.ToLower()

    # Convert the on-prem AD ObjectGUID to a Base64 string
    $immutableId = [System.Convert]::ToBase64String($user.ObjectGUID.ToByteArray())

    # Attempt to update
    try {
        # 1) Set OnPremisesImmutableId
        Update-MgUser -UserId $finalMatchedUPN -OnPremisesImmutableId $immutableId -ErrorAction Stop

        # 2) Mark extensionAttribute13 = "ScriptSynced"
        Update-MgUser -UserId $finalMatchedUPN -OnPremisesExtensionAttributes @{
            ExtensionAttribute13 = "ScriptSynced"
        } -ErrorAction Stop

        $Results += [pscustomobject]@{
            OnPremUser         = $user.UserPrincipalName
            AzureADUser        = $finalMatchedUPN
            Status             = "SUCCESS"
            Reason             = ""
            OnPremEnabled      = $user.Enabled
            NewID              = $immutableId
            PreviousID         = $azureUser.OnPremisesImmutableId
            DomainDiscrepancy  = $domainDiscrepancy
        }
        
        Write-Host "[ $currentUserIndex / $totalUsers ] SUCCESS: Set OnPremisesImmutableId for $($user.UserPrincipalName) => $finalMatchedUPN"
    }
    catch {
        $Results += [pscustomobject]@{
            OnPremUser         = $user.UserPrincipalName
            AzureADUser        = $finalMatchedUPN
            Status             = "FAILED"
            Reason             = $_.Exception.Message
            OnPremEnabled      = $user.Enabled
            NewID              = $immutableId
            PreviousID         = $azureUser.OnPremisesImmutableId
            DomainDiscrepancy  = $domainDiscrepancy
        }
        Write-Warning "[ $currentUserIndex / $totalUsers ] FAILED: $($user.UserPrincipalName) => $finalMatchedUPN : $($_.Exception.Message)"
    }
}

###############################################################################
# 11. Cross-reference: Identify Azure users not present in on-prem AD
###############################################################################
$OrphanedAzureUsers = @()
foreach ($key in $AzureUsersByUPN.Keys) {
    if (-not ($ProcessedUPNs -contains $key)) {
        $OrphanedAzureUsers += $AzureUsersByUPN[$key]
    }
}

if ($OrphanedAzureUsers.Count -gt 0) {
    Write-Host "`nFound $($OrphanedAzureUsers.Count) Azure AD user(s) that were NOT present in on-prem AD."
    
    if ($domainDiscrepancy) {
        "DomainDiscrepancy: $domainDiscrepancy" | Out-File $OrphanReportCsvPath
        $OrphanedAzureUsers |
            Select-Object `
                GivenName,
                @{Name="OnPremImmutableId"; Expression = { $_.OnPremisesImmutableId }},
                Surname,
                @{Name="UPN"; Expression = { $_.UserPrincipalName }} |
            Export-Csv -Path $OrphanReportCsvPath -NoTypeInformation -Append
    }
    else {
        $OrphanedAzureUsers |
            Select-Object `
                GivenName,
                @{Name="OnPremImmutableId"; Expression = { $_.OnPremisesImmutableId }},
                Surname,
                @{Name="UPN"; Expression = { $_.UserPrincipalName }} |
            Export-Csv -Path $OrphanReportCsvPath -NoTypeInformation
    }
    
    Write-Host "A CSV report of these users has been saved to: $OrphanReportCsvPath"
}
else {
    Write-Host "`nNo Azure AD users were found that are missing in on-prem AD."
}

###############################################################################
# 12. Output and export final results
###############################################################################
Write-Host "`n--- Summary of Hard Matching Operations ---"
$Results | Format-Table -AutoSize

if ($domainDiscrepancy) {
    "DomainDiscrepancy: $domainDiscrepancy" | Out-File $SuccessCsvPath
    $Results | Export-Csv -Path $SuccessCsvPath -NoTypeInformation -Append
}
else {
    $Results | Export-Csv -Path $SuccessCsvPath -NoTypeInformation
}
Write-Host "`nA CSV of all operations was saved to: $SuccessCsvPath"

$Failures = $Results | Where-Object { $_.Status -eq "FAILED" }
if ($Failures) {
    if ($domainDiscrepancy) {
        "DomainDiscrepancy: $domainDiscrepancy" | Out-File $FailureCsvPath
        $Failures | Export-Csv -Path $FailureCsvPath -NoTypeInformation -Append
    }
    else {
        $Failures | Export-Csv -Path $FailureCsvPath -NoTypeInformation
    }
    Write-Host "`nA CSV of all failures was saved to: $FailureCsvPath"
}
else {
    Write-Host "`nNo failures found - no CSV created."
}

###############################################################################
# 13. Disconnect from Microsoft Graph & Final Timer
###############################################################################
$elapsed = (Get-Date) - $scriptStart
Write-Host "Script completed successfully in $elapsed"

Disconnect-MgGraph | Out-Null
exit 0