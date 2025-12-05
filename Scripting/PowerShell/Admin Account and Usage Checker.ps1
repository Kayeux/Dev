param (
    [switch]$VerboseOutput
)

function Get-AdministratorsFromNetCommand {
    $adminMembers = @()
    $admins = net localgroup administrators | Select-Object -Skip 6 | Where-Object { $_ -match '\w' -and $_ -notmatch 'The command completed successfully' } | ForEach-Object {
        $name = $_.Trim()
        $type = if ($name -match '\\') { "Domain" } else { "Local" }

        $adminObject = New-Object PSObject -Property @{
            Name = $name
            Type = $type
        }

        $adminMembers += $adminObject
    }
    return $adminMembers
}

function Get-FrequentLogonUsers ($eventId) {
    $nameIndex = switch ($eventId) {
        4624 { 5 }   # For successful logons
        4673 { 1 }   # For privilege use
        default { 5 } # Default to the common index for logons
    }

    $logonEvents = Get-WinEvent -LogName Security -FilterXPath "*[System[EventID=$eventId]]"
    $logonUsers = $logonEvents | Where-Object { $_.Properties[$nameIndex].Value -ne 'SYSTEM' } |
                  Group-Object -Property { $_.Properties[$nameIndex].Value.Trim() } |
                  Sort-Object Count -Descending |
                  Select-Object Name, Count
    return $logonUsers
}

function CrossReferenceAdminsWithFrequentUsers ($admins, $frequentUsers) {
    $matches = @()
    foreach ($admin in $admins) {
        $adminNameNormalized = ($admin.Name.Split('\')[-1]).ToLower().Replace(" ", "")

        foreach ($user in $frequentUsers) {
            $userNameNormalized = ($user.Name.Split('@')[0]).ToLower().Replace(" ", "").TrimEnd('$')

            if ($userNameNormalized -eq $adminNameNormalized) {
                $matches += [PSCustomObject]@{
                    Name  = $admin.Name
                    Count = $user.Count
                    Type  = $admin.Type
                }
            }
        }
    }
    return $matches
}

# Main Script Execution
$administrators = Get-AdministratorsFromNetCommand

if ($VerboseOutput) {
    "Administrators Found:"
    $administrators | Format-Table -AutoSize
}

# Check for frequent logon users from event 4624
$frequentLogonUsers = Get-FrequentLogonUsers -eventId 4624
$matches = CrossReferenceAdminsWithFrequentUsers -admins $administrators -frequentUsers $frequentLogonUsers

if ($matches.Count -eq 0) {
    if ($VerboseOutput) {
        "No results found for Event 4624, checking Event 4673..."
    }
    # If no frequent logon users are found, check for event 4673
    $frequentLogonUsers = Get-FrequentLogonUsers -eventId 4673
}

if ($VerboseOutput -and $frequentLogonUsers.Count -gt 0) {
    "Frequent Logon Users Found:"
    $frequentLogonUsers | Format-Table -AutoSize
}

# Cross-reference admins with frequent logon users
$matches = CrossReferenceAdminsWithFrequentUsers -admins $administrators -frequentUsers $frequentLogonUsers

# Output Results for Matching Frequent Login Administrators
if ($matches.Count -gt 0) {
    Ninja-Property-Set userAsAdminFound $matches
    if ($VerboseOutput) {
        "Frequent Logon Administrators Found:"
        $matches | Format-Table -AutoSize
    }
} else {
    "No Frequent Logon Administrators Found."
}