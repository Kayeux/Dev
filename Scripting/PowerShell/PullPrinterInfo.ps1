# Datto Printer Inventory (Readable)
# - Prints a legible ASCII table to StdOut
# - Exports full data to CSV on disk
# - Default behavior: show network/shared printers only (filters out virtual/local like PDF/OneNote)
# References:
#   - Get-Printer / Get-PrinterPort are in PrintManagement :contentReference[oaicite:1]{index=1}
#   - Win32_TCPIPPrinterPort provides HostAddress for TCP/IP ports :contentReference[oaicite:2]{index=2}

param(
    [switch]$IncludeLocal,       # include local/virtual printers in output
    [switch]$LoadAllUserHives,   # best-effort: load offline user hives to find per-user printer connections
    [switch]$Detailed            # also print per-printer details (Format-List) after the table
)

$ErrorActionPreference = "SilentlyContinue"

$ComputerName = $env:COMPUTERNAME
$NowStamp     = (Get-Date).ToString("yyyyMMdd_HHmmss")
$OutDir       = Join-Path $env:ProgramData "DattoRMM\PrinterInventory"
New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
$CsvPath      = Join-Path $OutDir ("{0}_Printers_{1}.csv" -f $ComputerName, $NowStamp)

function Test-IsIPv4 {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -notmatch '^(?:\d{1,3}\.){3}\d{1,3}$') { return $false }
    $octets = $Value.Split('.')
    foreach ($o in $octets) {
        $n = 0
        if (-not [int]::TryParse($o, [ref]$n)) { return $false }
        if ($n -lt 0 -or $n -gt 255) { return $false }
    }
    return $true
}

function Ensure-HKUDrive {
    if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
        try { New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null } catch {}
    }
}

function Get-RegistryPortHint {
    param([string]$PortName)

    $candidates = @(
        "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports\$PortName",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Ports\$PortName",
        "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Ports\$PortName"
    )

    foreach ($p in $candidates) {
        try {
            if (Test-Path $p) {
                $props = Get-ItemProperty -Path $p
                $ip   = $null
                $hostname = $null

                if ($props.PSObject.Properties.Name -contains "IPAddress")          { $ip = [string]$props.IPAddress }
                if ($props.PSObject.Properties.Name -contains "PrinterHostAddress") { $ip = [string]$props.PrinterHostAddress }
                if ($props.PSObject.Properties.Name -contains "HostName")           { $hostname = [string]$props.HostName }

                if ($ip -or $hostname) {
                    return [pscustomobject]@{ IPAddress = $ip; HostName = $hostname; Source = "Registry" }
                }
            }
        } catch {}
    }
    return $null
}

function Is-NoisePrinter {
    param([pscustomobject]$Row)

    # If you *want* local/virtual printers, skip filtering
    if ($IncludeLocal) { return $false }

    $pn = [string]$Row.PrinterName
    $pt = [string]$Row.PortType
    $po = [string]$Row.PortName
    $dr = [string]$Row.DriverName

    # obvious virtual/local ports
    if ($pt -eq "LocalPort") { return $true }
    if ($po -match '^(PORTPROMPT:|nul:|FILE:|XPSPort:|FAX:|USB|LPT|COM)') { return $true }

    # common built-in virtual printers
    if ($pn -match 'Microsoft Print to PDF|Microsoft XPS Document Writer|OneNote|Fax') { return $true }
    if ($dr -match 'Microsoft Print To PDF|OneNote|XPS|Fax') { return $true }

    return $false
}

function Write-AsciiTable {
    param(
        [Parameter(Mandatory=$true)] [object[]]$Rows,
        [Parameter(Mandatory=$true)] [string[]]$Columns,
        [hashtable]$MaxWidth = @{}
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Output "(no rows)"
        return
    }

    # Sanitize + compute widths
    $width = @{}
    foreach ($c in $Columns) { $width[$c] = $c.Length }

    foreach ($r in $Rows) {
        foreach ($c in $Columns) {
            $v = ""
            try { $v = [string]$r.$c } catch { $v = "" }
            if ($null -eq $v) { $v = "" }
            $v = ($v -replace '[\r\n\t]+',' ').Trim()
            $cap = $null
            if ($MaxWidth.ContainsKey($c)) { $cap = [int]$MaxWidth[$c] }
            if ($cap -and $v.Length -gt $cap) { $v = $v.Substring(0,$cap-3) + "..." }
            if ($v.Length -gt $width[$c]) { $width[$c] = $v.Length }
        }
    }

    $line = "+"
    foreach ($c in $Columns) { $line += ("-" * ($width[$c] + 2)) + "+" }

    # Header
    Write-Output $line
    $hdr = "|"
    foreach ($c in $Columns) { $hdr += (" " + $c.PadRight($width[$c]) + " |") }
    Write-Output $hdr
    Write-Output $line

    # Rows
    foreach ($r in $Rows) {
        $row = "|"
        foreach ($c in $Columns) {
            $v = ""
            try { $v = [string]$r.$c } catch { $v = "" }
            if ($null -eq $v) { $v = "" }
            $v = ($v -replace '[\r\n\t]+',' ').Trim()
            $cap = $null
            if ($MaxWidth.ContainsKey($c)) { $cap = [int]$MaxWidth[$c] }
            if ($cap -and $v.Length -gt $cap) { $v = $v.Substring(0,$cap-3) + "..." }
            $row += (" " + $v.PadRight($width[$c]) + " |")
        }
        Write-Output $row
    }
    Write-Output $line
}

# Build port maps
$TcpIpPortMap = @{}
try {
    foreach ($p in (Get-CimInstance -ClassName Win32_TCPIPPrinterPort)) {
        if ($p.Name -and -not $TcpIpPortMap.ContainsKey($p.Name)) { $TcpIpPortMap[$p.Name] = $p }
    }
} catch {}

$PmPortMap = @{}
$HasPrintMgmt = $false
try {
    if (Get-Command Get-PrinterPort -ErrorAction SilentlyContinue) {
        $HasPrintMgmt = $true
        foreach ($pp in (Get-PrinterPort)) {
            if ($pp.Name -and -not $PmPortMap.ContainsKey($pp.Name)) { $PmPortMap[$pp.Name] = $pp }
        }
    }
} catch {}

# Collect printers
$Raw = @()

try {
    if (Get-Command Get-Printer -ErrorAction SilentlyContinue) {
        foreach ($pr in (Get-Printer)) {
            $Raw += [pscustomobject]@{
                PrinterName = [string]$pr.Name
                DriverName  = [string]$pr.DriverName
                PortName    = [string]$pr.PortName
                IsDefault   = [bool]$pr.Default
                IsShared    = [bool]$pr.Shared
                SharePath   = $null
                Source      = "Get-Printer"
                UserSid     = $null
            }
        }
    }
} catch {}

try {
    foreach ($wp in (Get-CimInstance -ClassName Win32_Printer)) {
        $share = $null
        if ($wp.Network -and $wp.Name -match '^\\\\') { $share = [string]$wp.Name }
        $Raw += [pscustomobject]@{
            PrinterName = [string]$wp.Name
            DriverName  = [string]$wp.DriverName
            PortName    = [string]$wp.PortName
            IsDefault   = [bool]$wp.Default
            IsShared    = [bool]$wp.Shared
            SharePath   = $share
            Source      = "Win32_Printer"
            UserSid     = $null
        }
    }
} catch {}

# Registry-based printer connections (HKLM + HKU loaded + optional offline user hives)
Ensure-HKUDrive
$LoadedMounts = @()

function Add-ConnectionKeys {
    param([string]$BasePath, [string]$SourceName, [string]$UserSid)

    try {
        if (-not (Test-Path $BasePath)) { return }
        foreach ($k in (Get-ChildItem -Path $BasePath)) {
            $child = [string]$k.PSChildName   # often ",,server,printer"
            $server = $null; $queue = $null; $sharePath = $null

            $parts = $child.Split(",")
            if ($parts.Length -ge 4) {
                $server = $parts[2]
                if ($parts.Length -gt 4) { $queue = ($parts[3..($parts.Length-1)] -join ",") } else { $queue = $parts[3] }
                if ($server -and $queue) { $sharePath = ("\\{0}\{1}" -f $server, $queue) }
            }

            $Raw += [pscustomobject]@{
                PrinterName = [string]($(if ($queue) { $queue } else { $child }))
                DriverName  = $null
                PortName    = [string]($(if ($sharePath) { $sharePath } else { $child }))
                IsDefault   = $false
                IsShared    = $true
                SharePath   = $sharePath
                Source      = $SourceName
                UserSid     = $UserSid
            }
        }
    } catch {}
}

Add-ConnectionKeys -BasePath "Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Print\Connections" -SourceName "HKLM_Connections" -UserSid $null

try {
    foreach ($sid in (Get-ChildItem HKU:\ | Where-Object { $_.PSChildName -match '^S-1-5-21-' })) {
        Add-ConnectionKeys -BasePath ("HKU:\{0}\Printers\Connections" -f $sid.PSChildName) -SourceName "HKU_Connections" -UserSid $sid.PSChildName
    }
} catch {}

if ($LoadAllUserHives) {
    # Best-effort: load NTUSER.DAT for profiles not currently loaded
    try {
        $profileList = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        foreach ($k in (Get-ChildItem $profileList)) {
            $sid = $k.PSChildName
            if ($sid -notmatch '^S-1-5-21-') { continue }

            $alreadyLoaded = Test-Path ("Registry::HKEY_USERS\{0}" -f $sid)
            if ($alreadyLoaded) { continue }

            $p = Get-ItemProperty -Path $k.PSPath
            $path = [string]$p.ProfileImagePath
            if ([string]::IsNullOrWhiteSpace($path)) { continue }

            $ntuser = Join-Path $path "NTUSER.DAT"
            if (-not (Test-Path $ntuser)) { continue }

            $mount = ("DATTO_TMP_{0}" -f ($sid -replace '[^0-9A-Za-z]','_'))
            $loadKey = ("HKU\{0}" -f $mount)

            # reg.exe is native + reliable for hive load/unload
            $rc = & reg.exe load $loadKey $ntuser 2>$null
            if ($LASTEXITCODE -eq 0) {
                $LoadedMounts += $mount
                Add-ConnectionKeys -BasePath ("HKU:\{0}\Printers\Connections" -f $mount) -SourceName "HKU_Connections(LoadedHive)" -UserSid $sid
            }
        }
    } catch {}
}

# Resolve + de-dup
$Seen = @{}
$Out  = @()

foreach ($r in $Raw) {
    $name  = [string]$r.PrinterName
    if ([string]::IsNullOrWhiteSpace($name)) { continue }

    $port  = [string]$r.PortName
    $share = [string]$r.SharePath
    $key   = ($name + "|" + $port + "|" + $share)
    if ($Seen.ContainsKey($key)) { continue }
    $Seen[$key] = $true

    $type = "Unknown"
    $ipOrHost = $null
    $portDetailSource = $null

    if ($port -match '^\\\\') {
        $type = "Shared"
        $portDetailSource = "ClientOnly"
    }
    elseif (Test-IsIPv4 $port) {
        $type = "DirectIP"
        $ipOrHost = $port
        $portDetailSource = "PortName"
    }
    elseif ($port -match '^IP_') {
        $maybe = $port.Substring(3)
        if (Test-IsIPv4 $maybe) {
            $type = "DirectIP"
            $ipOrHost = $maybe
            $portDetailSource = "PortName"
        }
    }

    if (-not $ipOrHost -and $port) {
        if ($TcpIpPortMap.ContainsKey($port)) {
            $type = "TCPIPPort"
            $ipOrHost = [string]$TcpIpPortMap[$port].HostAddress
            $portDetailSource = "Win32_TCPIPPrinterPort"
        }
        elseif ($HasPrintMgmt -and $PmPortMap.ContainsKey($port)) {
            # Only trust PrinterHostAddress here; do NOT print "Local Port" etc.
            $type = "PrinterPort"
            $ipOrHost = [string]$PmPortMap[$port].PrinterHostAddress
            $portDetailSource = "Get-PrinterPort"
        }
        else {
            $hint = Get-RegistryPortHint -PortName $port
            if ($hint) {
                if ($port -match '^WSD-') { $type = "WSD" }
                elseif ($port -match 'IPP|http') { $type = "IPP" }
                else { $type = "Port" }

                if ($hint.IPAddress) { $ipOrHost = [string]$hint.IPAddress }
                elseif ($hint.HostName) { $ipOrHost = [string]$hint.HostName }

                $portDetailSource = [string]$hint.Source
            }
            else {
                if ($port -match '^WSD-') { $type = "WSD" }
                elseif ($port -match '^(LPT|COM|USB|FILE|PORTPROMPT)') { $type = "LocalPort" }
            }
        }
    }

    $Out += [pscustomobject]@{
        ComputerName    = $ComputerName
        PrinterName     = $name
        DriverName      = [string]$r.DriverName
        PortName        = $port
        SharePath       = $share
        PortType        = $type
        IPAddressOrHost = $ipOrHost
        IsDefault       = [bool]$r.IsDefault
        IsShared        = [bool]$r.IsShared
        Source          = [string]$r.Source
        UserSid         = [string]$r.UserSid
        CollectedAt     = (Get-Date).ToString("s")
    }
}

# Unload any hives we loaded
foreach ($m in $LoadedMounts) {
    try { & reg.exe unload ("HKU\{0}" -f $m) 2>$null | Out-Null } catch {}
}

# Export full dataset to CSV on disk (for later pull)
try { $Out | Sort-Object PrinterName, PortName | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 } catch {}

# Build human-readable view
$View = @()
foreach ($o in $Out) {
    $conn = "DIRECT"
    if ($o.SharePath -and $o.SharePath -match '^\\\\') { $conn = "SHARED" }

    $View += [pscustomobject]@{
        PrinterName = $o.PrinterName
        Conn        = $conn
        IPHost      = [string]$o.IPAddressOrHost
        SharePath   = [string]$o.SharePath
        PortName    = [string]$o.PortName
        Default     = [string]$o.IsDefault
    }
}

# Filter noise unless IncludeLocal was set
$Filtered = @()
foreach ($o in $Out) {
    if (-not (Is-NoisePrinter -Row $o)) { $Filtered += $o }
}
$FilteredView = @()
foreach ($o in $Filtered) {
    $conn = "DIRECT"
    if ($o.SharePath -and $o.SharePath -match '^\\\\') { $conn = "SHARED" }
    $FilteredView += [pscustomobject]@{
        PrinterName = $o.PrinterName
        Conn        = $conn
        IPHost      = [string]$o.IPAddressOrHost
        SharePath   = [string]$o.SharePath
        PortName    = [string]$o.PortName
        Default     = [string]$o.IsDefault
    }
}

# Output (readable)
Write-Output ""
Write-Output ("PRINTER INVENTORY (Readable)  Computer={0}  CollectedAt={1}" -f $ComputerName, (Get-Date).ToString("s"))
Write-Output ("Saved CSV: {0}" -f $CsvPath)
Write-Output ("RawCount={0}  ShownCount={1}  (Use -IncludeLocal to show virtual/local)" -f $Out.Count, $Filtered.Count)
Write-Output ""

if ($FilteredView.Count -eq 0) {
    Write-Output "No network/shared printers found in this execution context."
    Write-Output "If you EXPECT printers but see none, run with -LoadAllUserHives or run the job as the logged-in user (Datto often runs as SYSTEM)."
    Write-Output ""
} else {
    Write-AsciiTable -Rows ($FilteredView | Sort-Object PrinterName, Conn) `
        -Columns @("PrinterName","Conn","IPHost","SharePath","PortName","Default") `
        -MaxWidth @{ PrinterName=45; IPHost=32; SharePath=55; PortName=28 }
    Write-Output ""
}

if ($Detailed) {
    Write-Output "DETAILS (per printer):"
    foreach ($o in ($Filtered | Sort-Object PrinterName, PortName)) {
        Write-Output "----------------------------------------"
        $o | Format-List ComputerName,PrinterName,DriverName,SharePath,PortName,PortType,IPAddressOrHost,IsDefault,IsShared,Source,UserSid,CollectedAt
    }
    Write-Output "----------------------------------------"
}

exit 0