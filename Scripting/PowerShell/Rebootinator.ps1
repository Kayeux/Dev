<#
.SYNOPSIS
    Displays a security update prompt to the user and handles system reboot if necessary.

.DESCRIPTION
    This script checks for pending updates and categorizes them into detailed groups.
    It prompts the user to reboot now, later, or schedule a reboot based on the execution count.
    Execution count warnings are displayed at specific thresholds.
    The script logs user interactions and scheduling details.

.PARAMETER EnableDebug
    Skips the uptime check for low uptimes.
    Outputs the displayed label to the console.
    Still increments execution count, shows standard label as needed, and functions as normal.

.PARAMETER PullData
    Only outputs the label data for uptime and updates required to the console.
    Does not increment execution count.
    Does not show dialog to the user.
    Skips uptime minimum check.

.PARAMETER TestUptime
    Skips all checks, only displays the "no updates required but you have a high uptime" dialog.
    Can still reboot from here.
    Does not increment execution count.

.EXAMPLE
    .\UpdatePrompt.ps1
    Displays the update prompt based on system uptime.

.EXAMPLE
    .\UpdatePrompt.ps1 -EnableDebug
    Skips the uptime check for low uptimes, outputs the displayed label to the console, and functions as normal.

.EXAMPLE
    .\UpdatePrompt.ps1 -PullData
    Only outputs the label data to the console without showing any dialogs.

.EXAMPLE
    .\UpdatePrompt.ps1 -TestUptime
    Displays the high uptime prompt regardless of pending updates or execution count.
#>

# Define parameters
param(
    [switch]$EnableDebug,
    [switch]$PullData,
    [switch]$TestUptime
)

# Configuration and Logging Paths
$ConfigDir = "$env:ProgramData\TechneauxCybersecurity"
$ConfigFile = "$ConfigDir\UpdateScriptConfig.json"
$LogDir = "$ConfigDir\Logs"
$LogFile = "$LogDir\UpdateScript.log"

# Ensure directories exist
foreach ($dir in @($ConfigDir, $LogDir)) {
    if (-not (Test-Path -Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

# Function to log messages
function Log-Message {
    param (
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")] [string]$Severity = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Severity] - $Message" | Out-File -FilePath $LogFile -Append -Encoding UTF8
}

# Function to initialize the configuration file
function Initialize-Config {
    if (-not (Test-Path -Path $ConfigFile)) {
        $config = @{
            ExecutionCount = 0
            LastRun = $null
            RebootScheduled = $false
        }
        $config | ConvertTo-Json | Out-File -FilePath $ConfigFile -Encoding UTF8
        Log-Message "Configuration file initialized."
    }
}

# Initialize configuration
Initialize-Config

# Function to read the configuration
function Get-Config {
    try {
        $configJson = Get-Content -Path $ConfigFile -Raw
        $config = $configJson | ConvertFrom-Json
        return $config
    }
    catch {
        if ($EnableDebug) {
            Write-Host "Error reading configuration: $_"
        }
        Log-Message "Error reading configuration: $_" "ERROR"
        return $null
    }
}

# Function to update the configuration
function Set-Config {
    param (
        $NewConfig
    )
    try {
        $NewConfig | ConvertTo-Json | Out-File -FilePath $ConfigFile -Encoding UTF8
        Log-Message "Configuration updated: ExecutionCount = $($NewConfig.ExecutionCount)"
    }
    catch {
        if ($EnableDebug) {
            Write-Host "Error writing configuration: $_"
        }
        Log-Message "Error writing configuration: $_" "ERROR"
    }
}

# Function to increment the execution count
function Increment-ExecutionCount {
    $config = Get-Config
    if ($config -ne $null) {
        $config.ExecutionCount += 1
        $config.LastRun = (Get-Date).ToString("o")
        Set-Config -NewConfig $config
    }
}

# Function to reset the execution count
function Reset-ExecutionCount {
    $config = Get-Config
    if ($config -ne $null) {
        $config.ExecutionCount = 0
        $config.LastRun = (Get-Date).ToString("o")
        $config.RebootScheduled = $false
        Set-Config -NewConfig $config
        Log-Message "Execution count reset to 0."
    }
}

# Define the maximum allowed executions before forcing a reboot
$MaxExecutions = 5

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get pending updates
function Get-PendingUpdates {
    try {
        if ($EnableDebug) { Write-Host "Initializing update session..." }
        # Initialize the update session and searcher
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        if ($updateSession -eq $null) {
            if ($EnableDebug) { Write-Host "Failed to create update session." }
            throw "Update session is null. Unable to proceed."
        } else {
            if ($EnableDebug) { Write-Host "Update session created successfully." }
        }

        $updateSearcher = $updateSession.CreateUpdateSearcher()
        if ($updateSearcher -eq $null) {
            if ($EnableDebug) { Write-Host "Failed to create update searcher." }
            throw "Update searcher is null. Unable to proceed."
        } else {
            if ($EnableDebug) { Write-Host "Update searcher created successfully." }
        }

        if ($EnableDebug) { Write-Host "Searching for pending updates..." }
        Log-Message "Searching for pending updates..."

        # Perform the search
        $searchCriteria = "IsInstalled=0 and IsHidden=0"
        if ($EnableDebug) { Write-Host "Search criteria: $searchCriteria" }
        $searchResult = $updateSearcher.Search($searchCriteria)
        if ($searchResult -eq $null) {
            if ($EnableDebug) { Write-Host "Search result is null." }
            throw "Update search result is null. Unable to retrieve updates."
        } else {
            if ($EnableDebug) { Write-Host "Search result obtained." }
        }

        # Get the updates collection
        $updates = $searchResult.Updates
        if ($updates -eq $null) {
            if ($EnableDebug) { Write-Host "Updates collection is null." }
            throw "Updates collection is null. Unable to retrieve updates."
        } else {
            if ($EnableDebug) { Write-Host "Updates collection obtained. Number of updates found: $($updates.Count)" }
        }

        # Simplified categories
        $updateCategories = @{
            "Critical Updates" = 0
            "Security Updates" = 0
            "Driver Updates" = 0
            "Feature Packs" = 0
            "Service Packs" = 0
            "Update Rollups" = 0
            "Definition Updates" = 0
            "Other Updates" = 0
        }

        $totalPending = $updates.Count
        if ($EnableDebug) { Write-Host "Total pending updates: $totalPending" }

        # Iterate through each update
        foreach ($update in $updates) {
            if ($EnableDebug) { Write-Host "Processing update: $($update.Title)" }
            $categorized = $false

            # Attempt to categorize using categories
            if ($update.Categories -ne $null) {
                if ($EnableDebug) { Write-Host " - Update.Categories count: $($update.Categories.Count)" }
                foreach ($category in $update.Categories) {
                    if ($category -eq $null) { continue }

                    $categoryTitle = $category.Title
                    if ($categoryTitle -eq $null) { $categoryTitle = "" }

                    $categoryTitle = $categoryTitle.Trim()
                    if ($EnableDebug) { Write-Host " - Category.Title: $categoryTitle" }

                    switch ($categoryTitle) {
                        "Critical Updates" {
                            $updateCategories["Critical Updates"] += 1
                            $categorized = $true
                        }
                        "Security Updates" {
                            $updateCategories["Security Updates"] += 1
                            $categorized = $true
                        }
                        "Drivers" {
                            $updateCategories["Driver Updates"] += 1
                            $categorized = $true
                        }
                        "Firmware" {
                            $updateCategories["Driver Updates"] += 1
                            $categorized = $true
                        }
                        "Definition Updates" {
                            $updateCategories["Definition Updates"] += 1
                            $categorized = $true
                        }
                        "Feature Packs" {
                            $updateCategories["Feature Packs"] += 1
                            $categorized = $true
                        }
                        "Service Packs" {
                            $updateCategories["Service Packs"] += 1
                            $categorized = $true
                        }
                        "Update Rollups" {
                            $updateCategories["Update Rollups"] += 1
                            $categorized = $true
                        }
                        default {
                            # Do nothing here
                        }
                    }
                }
            } else {
                if ($EnableDebug) { Write-Host " - Update.Categories is null." }
            }

            # Additional checks if not categorized yet
            if (-not $categorized) {
                # Check the DriverClass property
                if ($update.DriverClass -ne $null -and $update.DriverClass -ne "") {
                    if ($EnableDebug) { Write-Host " - DriverClass: $($update.DriverClass)" }
                    $updateCategories["Driver Updates"] += 1
                    $categorized = $true
                    if ($EnableDebug) { Write-Host " - Update categorized as 'Driver Updates' based on DriverClass property." }
                }
                # Fallback to Description and Title keywords
                elseif ($update.Description -match "(driver|firmware|security|critical|feature pack|service pack|rollup)") {
                    if ($EnableDebug) { Write-Host " - Matched keyword in Description." }
                    # Determine category based on keyword
                    if ($update.Description -match "driver|firmware") {
                        $updateCategories["Driver Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Driver Updates' based on Description." }
                    }
                    elseif ($update.Description -match "security|critical") {
                        $updateCategories["Security Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Security Updates' based on Description." }
                    }
                    elseif ($update.Description -match "feature pack") {
                        $updateCategories["Feature Packs"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Feature Packs' based on Description." }
                    }
                    elseif ($update.Description -match "service pack") {
                        $updateCategories["Service Packs"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Service Packs' based on Description." }
                    }
                    elseif ($update.Description -match "rollup") {
                        $updateCategories["Update Rollups"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Update Rollups' based on Description." }
                    }
                    else {
                        $updateCategories["Other Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Other Updates' based on Description." }
                    }
                    $categorized = $true
                }
                elseif ($update.Title -match "(driver|firmware|security|critical|feature pack|service pack|rollup)") {
                    if ($EnableDebug) { Write-Host " - Matched keyword in Title." }
                    # Determine category based on keyword
                    if ($update.Title -match "driver|firmware") {
                        $updateCategories["Driver Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Driver Updates' based on Title." }
                    }
                    elseif ($update.Title -match "security|critical") {
                        $updateCategories["Security Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Security Updates' based on Title." }
                    }
                    elseif ($update.Title -match "feature pack") {
                        $updateCategories["Feature Packs"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Feature Packs' based on Title." }
                    }
                    elseif ($update.Title -match "service pack") {
                        $updateCategories["Service Packs"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Service Packs' based on Title." }
                    }
                    elseif ($update.Title -match "rollup") {
                        $updateCategories["Update Rollups"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Update Rollups' based on Title." }
                    }
                    else {
                        $updateCategories["Other Updates"] += 1
                        if ($EnableDebug) { Write-Host " - Update categorized as 'Other Updates' based on Title." }
                    }
                    $categorized = $true
                }
                else {
                    $updateCategories["Other Updates"] += 1
                    if ($EnableDebug) { Write-Host " - Update categorized as 'Other Updates'" }
                }
            }
        }

        if ($EnableDebug) {
            Write-Host "Update categories and counts:"
            foreach ($key in $updateCategories.Keys) {
                Write-Host " - ${key}: $($updateCategories[$key])"
            }
        }

        return @{
            Total = $totalPending
            Categories = $updateCategories
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ($EnableDebug) { Write-Host "Failed to retrieve pending updates: $errorMessage" }
        Log-Message "Failed to retrieve pending updates: $errorMessage" "ERROR"
        return @{
            Total = 0
            Categories = @{}
        }
    }
}

# Function to estimate update installation time
function Get-UpdateEstimate {
    param (
        [int]$TotalUpdates,
        [hashtable]$UpdateCategories
    )

    # Average time per update category (in minutes)
    $timeEstimates = @{
        "Critical Updates" = 15
        "Security Updates" = 12
        "Driver Updates" = 8
        "Feature Packs" = 20
        "Service Packs" = 30
        "Update Rollups" = 25
        "Definition Updates" = 5
        "Other Updates" = 10
    }

    $totalEstimatedTime = 0

    foreach ($category in $UpdateCategories.Keys) {
        $count = $UpdateCategories[$category]
        if ($count -gt 0) {
            $timePerUpdate = $timeEstimates[$category]
            $totalEstimatedTime += $count * $timePerUpdate
        }
    }

    return $totalEstimatedTime
}

# Define category display names
$categoryDisplayNames = @{
    "Critical Updates" = "Critical Updates"
    "Security Updates" = "Security Updates"
    "Driver Updates" = "Driver Updates"
    "Feature Packs" = "Feature Packs"
    "Service Packs" = "Service Packs"
    "Update Rollups" = "Update Rollups"
    "Definition Updates" = "Definition Updates"
    "Other Updates" = "Other Updates"
}

# Get pending updates
$pendingUpdatesData = Get-PendingUpdates

# Get estimated update time
if ($pendingUpdatesData.Total -gt 0) {
    $estimatedTime = Get-UpdateEstimate -TotalUpdates $pendingUpdatesData.Total -UpdateCategories $pendingUpdatesData.Categories
    $estimatedTimeFormatted = "{0} minutes" -f $estimatedTime
} else {
    $estimatedTimeFormatted = "N/A"
}

# Get system uptime
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$lastBootUpTime = $os.LastBootUpTime
$uptime = (Get-Date) - $lastBootUpTime
$uptimeFormatted = "{0} days, {1} hours, {2} minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes

# Get configuration
$config = Get-Config

# Determine if we should increment execution count
$IncrementExecutionCount = -not ($PullData -or $TestUptime)

# Increment Execution Count if applicable
if ($IncrementExecutionCount) {
    Increment-ExecutionCount
    $config = Get-Config  # Refresh config
}

# Handle PullData parameter
if ($PullData) {
    # Build the message
    $message = "Machine Uptime: $uptimeFormatted`n"

    if ($pendingUpdatesData.Total -gt 0) {
        $message += "Pending Updates:`n"
        $message += " - Total: $($pendingUpdatesData.Total)`n"

        foreach ($category in $pendingUpdatesData.Categories.GetEnumerator()) {
            if ($category.Value -gt 0) {
                $displayName = $categoryDisplayNames[$category.Key]
                $message += " - ${displayName}: $($category.Value)`n"
            }
        }

        $message += "Estimated time to install updates: $estimatedTimeFormatted`n"
    } else {
        $message += "No pending updates.`n"
    }

    $message += "Execution Count: $($config.ExecutionCount) / $MaxExecutions`n"

    # Output message
    Write-Host $message

    # Do not show dialogs or proceed further
    return
}

# Determine if we should show the prompt
$SkipUptimeCheck = $EnableDebug -or $PullData -or $TestUptime
$showPrompt = $false
if ($SkipUptimeCheck -or $uptime.TotalDays -gt 3 -or $config.ExecutionCount -ge $MaxExecutions) {
    $showPrompt = $true
}

if (-not $showPrompt) {
    if ($EnableDebug) { Write-Host "Non Extensive Uptime" }
    Log-Message "Non Extensive Uptime"
    return
}

# Force Reboot Conditions
$forceReboot = $false
if ($config.ExecutionCount -ge $MaxExecutions) {
    $forceReboot = $true
}

# Function to show the scheduling form
function Show-SchedulingForm {
    # Scheduling Form
    $scheduleForm = New-Object System.Windows.Forms.Form
    $scheduleForm.Text = "Schedule Reboot"
    $scheduleForm.Size = New-Object System.Drawing.Size(350,250)
    $scheduleForm.StartPosition = "CenterScreen"
    $scheduleForm.TopMost = $true
    $scheduleForm.FormBorderStyle = "FixedDialog"
    $scheduleForm.MaximizeBox = $false
    $scheduleForm.MinimizeBox = $false
    $scheduleForm.ControlBox = $false

    $scheduleLabel = New-Object System.Windows.Forms.Label
    $scheduleLabel.Text = "Select reboot time (minutes from now):"
    $scheduleLabel.AutoSize = $false
    $scheduleLabel.Size = New-Object System.Drawing.Size(300,40)
    $scheduleLabel.Location = New-Object System.Drawing.Point(20,20)
    $scheduleLabel.TextAlign = "TopLeft"
    $scheduleLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $scheduleForm.Controls.Add($scheduleLabel)

    # Numeric UpDown control
    $numericUpDown = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDown.Minimum = 1
    $numericUpDown.Maximum = 1440  # Up to 24 hours
    $numericUpDown.Value = 5
    $numericUpDown.Location = New-Object System.Drawing.Point(20,70)
    $numericUpDown.Size = New-Object System.Drawing.Size(100,30)
    $numericUpDown.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $scheduleForm.Controls.Add($numericUpDown)

    # Expected Reboot Time Label
    $expectedTimeLabel = New-Object System.Windows.Forms.Label
    $expectedTimeLabel.Text = "Expected Reboot Time: "
    $expectedTimeLabel.AutoSize = $false
    $expectedTimeLabel.Size = New-Object System.Drawing.Size(300,40)
    $expectedTimeLabel.Location = New-Object System.Drawing.Point(20,110)
    $expectedTimeLabel.TextAlign = "TopLeft"
    $expectedTimeLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $scheduleForm.Controls.Add($expectedTimeLabel)

    # Update expected reboot time when value changes
    $numericUpDown.Add_ValueChanged({
        $delay = [int]$numericUpDown.Value
        $rebootTime = (Get-Date).AddMinutes($delay)
        $expectedTimeLabel.Text = "Expected Reboot Time: $($rebootTime.ToString('g'))"
    })

    # Initialize expected reboot time
    $delay = [int]$numericUpDown.Value
    $rebootTime = (Get-Date).AddMinutes($delay)
    $expectedTimeLabel.Text = "Expected Reboot Time: $($rebootTime.ToString('g'))"

    # Confirm Schedule Button
    $confirmScheduleButton = New-Object System.Windows.Forms.Button
    $confirmScheduleButton.Text = "Confirm"
    $confirmScheduleButton.Size = New-Object System.Drawing.Size(120,40)
    $confirmScheduleButton.Location = New-Object System.Drawing.Point(50,160)
    $confirmScheduleButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $confirmScheduleButton.Add_Click({
        $delay = [int]$numericUpDown.Value
        $rebootTime = (Get-Date).AddMinutes($delay)
        Log-Message "User scheduled a reboot in $delay minutes at $($rebootTime.ToString('g'))."
        Write-Host "Reboot scheduled in $delay minutes at $($rebootTime.ToString('g'))."
        shutdown.exe /r /t ($delay * 60)
        Reset-ExecutionCount
        $scheduleForm.Close()
        $form.Close()
    })
    $scheduleForm.Controls.Add($confirmScheduleButton)

    # Cancel Schedule Button
    $cancelScheduleButton = New-Object System.Windows.Forms.Button
    $cancelScheduleButton.Text = "Cancel"
    $cancelScheduleButton.Size = New-Object System.Drawing.Size(120,40)  # Match size
    $cancelScheduleButton.Location = New-Object System.Drawing.Point(180,160)
    $cancelScheduleButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $cancelScheduleButton.Add_Click({
        Write-Host "User canceled scheduling the reboot."
        Log-Message "User canceled scheduling the reboot."
        $scheduleForm.Close()
    })
    $scheduleForm.Controls.Add($cancelScheduleButton)

    # Show scheduling form
    $scheduleForm.TopMost = $true
    $scheduleForm.Add_Shown({$scheduleForm.Activate()})
    [void]$scheduleForm.ShowDialog()
}

# Build the message
$message = @"
Important updates or maintenance are required. Your system has not been properly rebooted for $uptimeFormatted.
"@

# Adjust message based on TestUptime parameter
if ($TestUptime) {
    $pendingUpdatesData.Total = 0
    $pendingUpdatesData.Categories = @{}
}

if ($pendingUpdatesData.Total -gt 0) {
    $message += "`nPending Updates:`n"
    $message += " - Total: $($pendingUpdatesData.Total)`n"

    foreach ($category in $pendingUpdatesData.Categories.GetEnumerator()) {
        if ($category.Value -gt 0) {
            $displayName = $categoryDisplayNames[$category.Key]
            $message += " - ${displayName}: $($category.Value)`n"
        }
    }

    $message += "`nEstimated time to install updates: $estimatedTimeFormatted`n"
} else {
    $message += "`nWhile there are no pending updates, your system has been running for an extended period which may affect performance. A reboot is recommended.`n"
}

$message += @"
Execution Count: $($config.ExecutionCount) / $MaxExecutions

A reboot is necessary to ensure your system's performance and security.
"@

# Output message to console if EnableDebug is set
if ($EnableDebug) {
    Write-Host $message
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Security Update Required -- Techneaux Cybersecurity Team"
$form.Size = New-Object System.Drawing.Size(560,450)  # Increased width by 10
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.ControlBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Text = $message
$label.AutoSize = $false
$label.Size = New-Object System.Drawing.Size(520,250)  # Increased width by 10
$label.Location = New-Object System.Drawing.Point(20,20)
$label.TextAlign = "TopLeft"
$label.Font = New-Object System.Drawing.Font("Segoe UI",10)
$form.Controls.Add($label)

# Yes Button
$yesButton = New-Object System.Windows.Forms.Button
$yesButton.Text = "Reboot Now"
$yesButton.Size = New-Object System.Drawing.Size(140,30)
$yesButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
if ($forceReboot) {
    $yesButton.Location = New-Object System.Drawing.Point(110,320)  # Adjusted for increased width
} else {
    $yesButton.Location = New-Object System.Drawing.Point(90,320)
}
$yesButton.Add_Click({
    $form.Hide()
    # Confirmation Dialog
    $confirmForm = New-Object System.Windows.Forms.Form
    $confirmForm.Text = "Confirm Reboot"
    $confirmForm.Size = New-Object System.Drawing.Size(400,200)
    $confirmForm.StartPosition = "CenterScreen"
    $confirmForm.TopMost = $true
    $confirmForm.FormBorderStyle = "FixedDialog"
    $confirmForm.MaximizeBox = $false
    $confirmForm.MinimizeBox = $false
    $confirmForm.ControlBox = $false

    $confirmMessage = "Please ensure that you have saved all your work and closed all applications. Do you wish to proceed with the reboot?"

    $confirmLabel = New-Object System.Windows.Forms.Label
    $confirmLabel.Text = $confirmMessage
    $confirmLabel.AutoSize = $false
    $confirmLabel.Size = New-Object System.Drawing.Size(360,80)
    $confirmLabel.Location = New-Object System.Drawing.Point(20,20)
    $confirmLabel.TextAlign = "MiddleLeft"
    $confirmLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $confirmForm.Controls.Add($confirmLabel)

    # Confirm Button
    $confirmButton = New-Object System.Windows.Forms.Button
    $confirmButton.Text = "Yes, Reboot"
    $confirmButton.Size = New-Object System.Drawing.Size(120,40)
    $confirmButton.Location = New-Object System.Drawing.Point(70,120)
    $confirmButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $confirmButton.Add_Click({
        Write-Host "User chose to reboot now."
        Log-Message "User chose to reboot now."
        shutdown.exe /r /t 60
        Write-Host "Reboot scheduled in 60 seconds."
        Log-Message "Reboot scheduled in 60 seconds."
        Reset-ExecutionCount
        $confirmForm.Close()
    })
    $confirmForm.Controls.Add($confirmButton)

    # Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Size = New-Object System.Drawing.Size(120,40)
    $cancelButton.Location = New-Object System.Drawing.Point(210,120)
    $cancelButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $cancelButton.Add_Click({
        Write-Host "User canceled the reboot."
        Log-Message "User canceled the reboot."
        $confirmForm.Close()
        $form.Show()
    })
    $confirmForm.Controls.Add($cancelButton)

    # Show confirmation dialog
    $confirmForm.TopMost = $true
    $confirmForm.Add_Shown({$confirmForm.Activate()})
    [void]$confirmForm.ShowDialog()
})
$form.Controls.Add($yesButton)

# No Button (only if not forced reboot at max count)
if (-not $forceReboot -or $config.ExecutionCount -eq ($MaxExecutions - 1)) {
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Text = "Later"
    $noButton.Size = New-Object System.Drawing.Size(140,30)
    $noButton.Location = New-Object System.Drawing.Point(310,320)  # Adjusted for increased width
    $noButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $noButton.Add_Click({
        Write-Host "User chose to reboot later."
        Log-Message "User chose to reboot later."

        # Replace the MessageBox with the custom scheduling prompt
        Show-SchedulePrompt
    })
    $form.Controls.Add($noButton)
}

# Schedule Button
$scheduleButton = New-Object System.Windows.Forms.Button
$scheduleButton.Text = "Schedule Reboot"
$scheduleButton.Size = New-Object System.Drawing.Size(140,30)
$scheduleButton.Location = New-Object System.Drawing.Point(200,370)  # Adjusted for increased width
$scheduleButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
$scheduleButton.Add_Click({
    Show-SchedulingForm
})
$form.Controls.Add($scheduleButton)

# If forced reboot at max count, remove "Later" button
if ($forceReboot -and $config.ExecutionCount -ge $MaxExecutions) {
    if ($noButton) {
        $form.Controls.Remove($noButton)
    }
    # Adjust positions of buttons
    $yesButton.Location = New-Object System.Drawing.Point(140,320)
    $scheduleButton.Location = New-Object System.Drawing.Point(310,320)
}

# Function to show the custom scheduling prompt with easter egg
function Show-SchedulePrompt {
    # Create a form to mimic the MessageBox
    $promptForm = New-Object System.Windows.Forms.Form
    $promptForm.Text = "Schedule Reboot"
    $promptForm.Size = New-Object System.Drawing.Size(400,200)
    $promptForm.StartPosition = "CenterScreen"
    $promptForm.TopMost = $true
    $promptForm.FormBorderStyle = "FixedDialog"
    $promptForm.MaximizeBox = $false
    $promptForm.MinimizeBox = $false
    $promptForm.ControlBox = $false

    # Label
    $promptLabel = New-Object System.Windows.Forms.Label
    $promptLabel.Text = "Would you like to schedule a reboot now?"
    $promptLabel.AutoSize = $false
    $promptLabel.Size = New-Object System.Drawing.Size(360,80)
    $promptLabel.Location = New-Object System.Drawing.Point(20,20)
    $promptLabel.TextAlign = "MiddleCenter"
    $promptLabel.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $promptLabel.TabStop = $false
    $promptForm.Controls.Add($promptLabel)

    # Yes Button
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Text = "Yes"
    $yesButton.Size = New-Object System.Drawing.Size(100,30)
    $yesButton.Location = New-Object System.Drawing.Point(90,120)
    $yesButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $yesButton.TabStop = $false
    $yesButton.Add_Click({
        # Open the scheduling form
        Show-SchedulingForm
        $promptForm.Close()
    })
    $promptForm.Controls.Add($yesButton)

    # No Button
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Text = "No"
    $noButton.Size = New-Object System.Drawing.Size(100,30)
    $noButton.Location = New-Object System.Drawing.Point(210,120)
    $noButton.Font = New-Object System.Drawing.Font("Segoe UI",10)
    $noButton.TabStop = $false
    $noButton.Add_Click({
        # User chose not to schedule, close the form
        $promptForm.Close()
        $form.Close()
    })
    $promptForm.Controls.Add($noButton)

    # Handle the PreviewKeyDown event for the buttons to capture arrow keys
    $yesButton.Add_PreviewKeyDown({
        param($sender, $e)
        $e.IsInputKey = $true
    })
    $noButton.Add_PreviewKeyDown({
        param($sender, $e)
        $e.IsInputKey = $true
    })

    # Easter Egg Key Sequence
    $keySequence = @(
        [System.Windows.Forms.Keys]::Up,
        [System.Windows.Forms.Keys]::Up,
        [System.Windows.Forms.Keys]::Down,
        [System.Windows.Forms.Keys]::Down,
        [System.Windows.Forms.Keys]::Left,
        [System.Windows.Forms.Keys]::Right,
        [System.Windows.Forms.Keys]::Left,
        [System.Windows.Forms.Keys]::Right,
        [System.Windows.Forms.Keys]::ControlKey,
        [System.Windows.Forms.Keys]::ShiftKey
    )
    $keyBuffer = New-Object System.Collections.Generic.Queue[System.Windows.Forms.Keys]

    # Enable Key Preview to capture key events
    $promptForm.KeyPreview = $true
    $promptForm.Add_KeyDown({
        param($sender, $e)

        # Enqueue the key
        $keyBuffer.Enqueue($e.KeyCode)

        # Trim the buffer to the length of the key sequence
        while ($keyBuffer.Count -gt $keySequence.Length) {
            $keyBuffer.Dequeue() | Out-Null
        }

        # Debugging output (remove or comment out after testing)
        # Write-Host "Key pressed: $($e.KeyCode)"
        # Write-Host "Current buffer: $($keyBuffer.ToArray() -join ', ')"

        # Check if the sequence matches
        $bufferArray = $keyBuffer.ToArray()
        $sequenceMatches = $true
        for ($i = 0; $i -lt $keySequence.Length; $i++) {
            if ($bufferArray[$i] -ne $keySequence[$i]) {
                $sequenceMatches = $false
                break
            }
        }

        if ($sequenceMatches -and $keyBuffer.Count -eq $keySequence.Length) {
            # Easter Egg Activated
            [System.Windows.Forms.MessageBox]::Show("Easter Egg Activated!", "Easter Egg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            # Open the browser to the specified URL
            Start-Process "https://www.youtube.com/watch?v=-8H4GKg-mYQ"
            # Clear the key buffer
            $keyBuffer.Clear()
        }
    })

    # Ensure the form itself has focus
    $promptForm.Add_Shown({
        $promptForm.Activate()
        $promptForm.Focus()
    })
    [void]$promptForm.ShowDialog()
}


# Show the main form
$form.TopMost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()

# **End of Script**