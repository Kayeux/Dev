# Run Get-SmbConnection to get all active SMB connections
$connections = Get-SmbConnection | Where-Object { $_.Dialect -notmatch "3.*" }

# Check if any connections are using SMB versions lower than SMB 3.0
if ($connections) {
    # Output the connections using inferior SMB versions
    Write-Output "The following connections are using SMB versions lower than SMB 3:"
    $connections | Select-Object ServerName, ShareName, Dialect | Format-Table -AutoSize
} else {
    # No connections found with SMB 1.0, 2.0, or 2.1 - proceed to disable SMB protocols
    Write-Output "No connections using SMB versions lower than SMB 3. Proceeding to disable SMB 1 and SMB 2."

    # Disable SMB 1 and SMB 2 protocols
    Set-SmbServerConfiguration -EnableSMB1Protocol $false -Confirm:$false
    Set-SmbServerConfiguration -EnableSMB2Protocol $false -Confirm:$false

    Write-Output "SMB 1 and SMB 2 protocols have been disabled."
}