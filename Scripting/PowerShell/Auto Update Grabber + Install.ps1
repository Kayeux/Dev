# Define a parameter for the URL of the update file
param(
    [string]$updateUrl
)

# Function to check if the update is already installed
function Is-UpdateInstalled {
    param (
        [string]$updateKbNumber
    )

    $installedUpdates = Get-HotFix | Where-Object { $_.HotFixID -eq $updateKbNumber }
    return $installedUpdates -ne $null
}

# Define the path to save the downloaded file in the Temp folder
$tempFolder = [System.IO.Path]::GetTempPath()
$tempFile = [System.IO.Path]::Combine($tempFolder, "update.msu")

# Extract KB number from the URL (assuming the URL contains the KB number)
$kbNumber = [System.Text.RegularExpressions.Regex]::Match($updateUrl, "KB\d+").Value

if ($kbNumber) {
    Write-Output "Checking if the update $kbNumber is already installed..."

    if (Is-UpdateInstalled -updateKbNumber $kbNumber) {
        Write-Output "Update $kbNumber is already installed. No action is required."
        exit
    } else {
        Write-Output "Update $kbNumber is not installed. Proceeding with download and installation..."

        # Download the update file using curl
        Write-Output "Downloading the update file from $updateUrl..."
        Invoke-Expression "curl.exe -o $tempFile $updateUrl"

        # Check if the file was downloaded successfully
        if (Test-Path $tempFile) {
            Write-Output "Downloaded file saved to $tempFile."

            # Execute the downloaded update file
            Write-Output "Executing the update file..."
            Start-Process "wusa.exe" -ArgumentList "$tempFile /quiet /norestart" -Wait

            Write-Output "Update installation initiated. Check Windows Update for progress."
        } else {
            Write-Output "Failed to download the update file. Please check the URL and try again."
        }
    }
} else {
    Write-Output "Could not extract KB number from the URL. Please ensure the URL is correct."
}
