# Run this script as Administrator

# Define the base paths to search
$programFilesPaths = @(
    "${env:ProgramFiles}",
    "${env:ProgramFiles(x86)}"
)

# Define possible directory names
$dirNames = @(
    'Intel(R) Rapid Storage Technology enterprise',
    'Intel(R) Rapid Storage Technology',
    'Intel Rapid Storage Technology'
)

# Initialize an array to store found paths
$foundPaths = @()

# Search for directories
foreach ($basePath in $programFilesPaths) {
    foreach ($dirName in $dirNames) {
        $intelDir = Join-Path -Path $basePath -ChildPath "Intel"
        $fullPath = Join-Path -Path $intelDir -ChildPath $dirName
        if (Test-Path $fullPath) {
            Write-Host "Found directory: $fullPath"
            $foundPaths += $fullPath
        }
    }
}

if ($foundPaths.Count -eq 0) {
    Write-Host "No Intel RST directories found."
    exit
}

# Process each found directory
foreach ($path in $foundPaths) {
    Write-Host "Processing directory: $path"

    # Take ownership of the directory and its contents
    takeown /F "$path" /R /D Y | Out-Null

    # Grant full control to Administrators and SYSTEM
    icacls "$path" /grant:r "Administrators:F" "SYSTEM:F" /T /C | Out-Null

    # Disable inheritance and remove all inherited permissions
    icacls "$path" /inheritance:r /T /C | Out-Null

    # Get all ACL entries
    $acl = Get-Acl -Path "$path"

    # Remove all access rules for users that are not Administrators or SYSTEM
    $acl.Access | ForEach-Object {
        if (
            $_.IdentityReference -notmatch '^(BUILTIN\\Administrators|NT AUTHORITY\\SYSTEM)$'
        ) {
            if ($_.IsInherited -eq $false) {
                Write-Host "Removing permission for $($_.IdentityReference) on $path"
                $acl.RemoveAccessRule($_)
            }
        }
    }

    # Set the updated ACL
    Set-Acl -Path "$path" -AclObject $acl

    # Repeat for all files and subdirectories
    Get-ChildItem -Path "$path" -Recurse -Force | ForEach-Object {
        # Get the ACL
        $itemAcl = Get-Acl -Path $_.FullName

        # Remove all access rules for users that are not Administrators or SYSTEM
        $itemAcl.Access | ForEach-Object {
            if (
                $_.IdentityReference -notmatch '^(BUILTIN\\Administrators|NT AUTHORITY\\SYSTEM)$'
            ) {
                if ($_.IsInherited -eq $false) {
                    Write-Host "Removing permission for $($_.IdentityReference) on $($_.FullName)"
                    $itemAcl.RemoveAccessRule($_)
                }
            }
        }

        # Set the updated ACL
        Set-Acl -Path $_.FullName -AclObject $itemAcl
    }
}

Write-Host "Permission adjustments completed."
