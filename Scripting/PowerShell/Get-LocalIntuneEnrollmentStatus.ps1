function Get-LocalIntuneEnrollmentStatus {
    <#
    .SYNOPSIS
        Checks whether this device is Intune-enrolled at the machine (HKLM) level.

    .DESCRIPTION
        Scans HKLM:\SOFTWARE\Microsoft\Enrollments for a subkey indicating Intune enrollment:
          - EnrollmentState = 1 (success)
          - ProviderID = "MS DM Server" or "Microsoft Device Management"

        Returns a custom PSObject with:
          - IsIntuneEnrolled: Boolean
          - Details: Any matching subkey name(s) or an explanation
    #>

    [CmdletBinding()]
    param()

    # Define results object
    $result = [PSCustomObject]@{
        IsIntuneEnrolled = $false
        Details          = "No device-level Intune enrollment subkey found."
    }

    $enrollPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
    if (-not (Test-Path $enrollPath)) {
        $result.Details = "HKLM:\SOFTWARE\Microsoft\Enrollments does not exist. No device MDM enrollment."
        return $result
    }

    $foundMatches = @()

    # Remove -Directory and filter for containers manually
    Get-ChildItem -Path $enrollPath -ErrorAction SilentlyContinue |
        Where-Object { $_.PSIsContainer } |
        ForEach-Object {
            try {
                $props = Get-ItemProperty -Path $_.PSPath -ErrorAction Stop
                $state = $props.EnrollmentState
                $provider = $props.ProviderID

                # Check for typical Intune patterns
                if (($state -gt 0) -and ($provider -match "(?i)(MS DM Server|Microsoft Device Management)")) {
                    $foundMatches += $_.PSChildName
                }
            }
            catch {
                # Just ignore read errors
            }
        }

    if ($foundMatches) {
        $result.IsIntuneEnrolled = $true
        $result.Details = "Found Intune enrollment subkey(s): $($foundMatches -join ', ')"
    }

    return $result
}
