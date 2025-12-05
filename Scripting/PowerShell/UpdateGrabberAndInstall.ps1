#This script expects an argument using the -updateUrl flag with a link being required within quotes "somelink\to\patch.msu" created with the intent to manually push patches that could resolve CVEs
#Example Call
#.\scriptName.ps1 -updateUrl "https://catalog.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/92ac27d7-6832-4bac-8205-922458d3df2b/public/windows11.0-kb5040527-x64_4713766dc272c376bee6d39d39a84a85bcd7f1e7.msu"
#Script handles the downloading and installation of the presented msu file

#NOTE: This script verifies if the patch is already installed by checking the KB#### of the file and cross referencing installed updates
#HOWEVER, this has presented issues when similar KB patches are installed with differing DisplayVersion ( XXHX ( 23H2 ) )
# Attempted to remediate this by utilizing an updateID instead of download link and then pulling info

# But going through the updateID method we were unable to properly find the download link and proceed ( because in Microsoft Catalog the download link opens 
# a new window using JS params instead of a window we can directly link to and I am unsure how to emulate a click without using tools like Selenium / CON IE

#In the event that a KB### patch from a previous DisplayVersion is installed, the script WILL fail but present as a success
# To proceed this would require an uninstallation of the older patch, but this may not be best practice

# We have a few modifiable values such as a Timeout detector which kills the script if elapsed time has passed
# This timeout detector presents itself as a pseudo download bar

# Define a parameter for the URL of the update file
param(
    [string]$updateUrl
)

# Function to check if the update is already installed
function Is-UpdateInstalled {
    param (
        [string]$updateKbNumber
    )

    $installedUpdates = Get-HotFix | Where-Object { $_.HotFixID -eq "KB"+$updateKbNumber }
    return $installedUpdates -ne $null
}

# Function to check if another installation is in progress
function Is-AnotherInstallationInProgress {
    $inProgress = Get-WmiObject -Query "SELECT * FROM Win32_Process WHERE Name='msiexec.exe'"
    return $inProgress -ne $null
}

# Function to install the update
function Install-Update {
    param (
        [string]$updateUrl
    )

    # Extract KB number from the URL (assuming the URL contains the KB number)
    $kbNumber = [System.Text.RegularExpressions.Regex]::Match($updateUrl, "kb\d+", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Value.ToUpper()

    if ($kbNumber) {
        Write-Output "Checking if the update $kbNumber is already installed..."

        if (Is-UpdateInstalled -updateKbNumber $kbNumber) {
            Write-Output "Update $kbNumber is already installed. No action is required."
            return $true
        } else {
            Write-Output "Update $kbNumber is not installed. Proceeding with download and installation..."

            # Define the path to save the downloaded file in the Temp folder
            $tempFolder = [System.IO.Path]::GetTempPath()
            $tempFile = [System.IO.Path]::Combine($tempFolder, "$kbNumber.msu")

            # Check if the update file already exists in the temp folder
            if (Test-Path $tempFile) {
                Write-Output "Update file already exists in the temp folder. Skipping download."
            } else {
                Write-Output "Downloading the update file from $updateUrl..."
                Invoke-Expression "curl.exe -o $tempFile $updateUrl"

                # Check if the file was downloaded successfully
                if (-not (Test-Path $tempFile)) {
                    Write-Output "Failed to download the update file. Please check the URL and try again."
                    return $false
                }
            }

            Write-Output "Downloaded file saved to $tempFile."

            # Execute the downloaded update file and capture the output
            Write-Output "Executing the update file..."
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = "wusa.exe"
            $processInfo.Arguments = "$tempFile /quiet /norestart"
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null
            $standardOutput = $process.StandardOutput
            $standardError = $process.StandardError

            $timeout = 720  # Timeout in seconds (12 minutes)
            $elapsedTime = 0
            $interval = 5  # Check interval in seconds
            $progressIncrement = [math]::Round((100 / ($timeout / $interval)))

            Write-Progress -Activity "Installing Update" -Status "In Progress" -PercentComplete 0

            while (!$process.HasExited -and $elapsedTime -lt $timeout) {
                Start-Sleep -Seconds $interval
                $elapsedTime += $interval
                $progressPercent = [math]::Round((($elapsedTime / $timeout) * 100), 0)
                Write-Progress -Activity "Installing Update" -Status "In Progress" -PercentComplete $progressPercent
                Write-Output "Update installation in progress... ($elapsedTime seconds elapsed)"
            }

            if (!$process.HasExited) {
                Write-Output "The update installation process timed out. Terminating the process."
                $process.Kill()
            }

            $exitCode = $process.ExitCode
            $output = $standardOutput.ReadToEnd()
            $errorOut = $standardError.ReadToEnd()

            Write-Output "WUSA Exit Code: $exitCode"
            Write-Output "WUSA Output: $output"
            Write-Output "WUSA Error: $errorOut"

            switch ($exitCode) {
                -2145124329 { Write-Output "Error 0x80240017: Unspecified error. The update could not be installed. This may indicate a problem with the update package or that the update is not applicable." }
                0xca00a009 { Write-Output "The update $kbNumber is not applicable to this machine." }
                1618 { Write-Output "Another installation is already in progress. Retrying..." }
                0 { Write-Output "Update installation completed successfully." }
                default { Write-Output "An unknown error occurred during the update installation. Exit Code: $exitCode" }
            }

            if ($exitCode -eq 0) {
                Write-Output "Checking if the update $kbNumber is now installed..."
                if (Is-UpdateInstalled -updateKbNumber $kbNumber) {
                    Write-Output "Update $kbNumber was successfully installed."
                    return $true
                } else {
                    Write-Output "Update $kbNumber installation failed or is not yet completed. Please check Windows Update for progress."
                    return $false
                }
            } elseif ($exitCode -eq 1618) {
                return $false  # Another installation is in progress, will retry
            } else {
                return $false
            }
        }
    } else {
        Write-Output "Could not extract KB number from the URL. Please ensure the URL is correct."
        return $false
    }
}

# Main script logic
$maxRetries = 3
$retryCount = 0
$retryInterval = 60  # Interval in seconds to wait before retrying

while ($retryCount -lt $maxRetries) {
    if (-not (Is-AnotherInstallationInProgress)) {
        if (Install-Update -updateUrl $updateUrl) {
            Write-Output "Update installation succeeded."
            exit
        } else {
            Write-Output "Update installation failed. Retrying..."
        }
    } else {
        Write-Output "Another installation is in progress. Waiting for $retryInterval seconds before retrying..."
        Start-Sleep -Seconds $retryInterval
    }
    $retryCount++
}

Write-Output "Update installation failed after $maxRetries attempts."
exit 1