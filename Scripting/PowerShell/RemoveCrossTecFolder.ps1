# Define the directory path
$directoryPath = "C:\ProgramData\crosstec"

# Check if the directory exists
if (Test-Path -Path $directoryPath) {
    Write-Output "Directory exists: $directoryPath"
    
    # Attempt to remove the directory
    try {
        Remove-Item -Path $directoryPath -Recurse -Force
        Write-Output "Directory '$directoryPath' has been successfully deleted."
    } catch {
        Write-Output "Failed to delete the directory '$directoryPath'. Error: $_"
    }
} else {
    Write-Output "Directory does not exist: $directoryPath"
}