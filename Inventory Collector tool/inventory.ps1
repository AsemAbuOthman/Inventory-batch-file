
$basePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputFolder = "$basePath\output"

if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

$time = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$output = "$outputFolder\FULL_IT_INVENTORY_$time.csv"

# ================= INPUT =================
$employeeName = Read-Host "Enter Employee Name"
$phone = Read-Host "Enter Phone Number (optional)"

$email = "$env:USERNAME@$env:USERDOMAIN.local"

# ================= SYSTEM =================
$cs = Get-CimInstance Win32_ComputerSystem
$os = Get-CimInstance Win32_OperatingSystem
$bios = Get-CimInstance Win32_BIOS
$board = Get-CimInstance Win32_BaseBoard

# ================= DEVICE TYPE (Laptop/Desktop) =================
$chassis = Get-CimInstance Win32_SystemEnclosure
$chassisType = ($chassis.ChassisTypes | Select-Object -First 1)

$deviceType = switch ($chassisType) {
    8 {"Laptop"}
    9 {"Laptop"}
    10 {"Laptop"}
    11 {"Laptop"}
    12 {"Laptop Docked"}
    3 {"Desktop"}
    4 {"Desktop"}
    5 {"Desktop"}
    default {"Unknown"}
}

# ================= CPU / GPU =================
$cpu = Get-CimInstance Win32_Processor
$gpu = Get-CimInstance Win32_VideoController

# ================= RAM =================
$ram = Get-CimInstance Win32_PhysicalMemory

$ramDetails = ($ram | ForEach-Object {
    $type = switch ($_.SMBIOSMemoryType) {
        20 {"DDR"}
        21 {"DDR2"}
        24 {"DDR3"}
        26 {"DDR4"}
        34 {"DDR5"}
        default {"Unknown"}
    }

    "$type $([math]::Round($_.Capacity/1GB,2))GB @ $($_.Speed)MHz"
}) -join " | "

$ramTotal = [math]::Round(($ram | Measure-Object Capacity -Sum).Sum / 1GB,2)

# ================= DISKS =================
$disks = Get-CimInstance Win32_DiskDrive
$logical = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"

$diskDetails = ($disks | ForEach-Object {
    "$($_.Model) - $([math]::Round($_.Size/1GB))GB"
}) -join " | "

$driveUsage = ($logical | ForEach-Object {
    $size = [math]::Round($_.Size/1GB,1)
    $free = [math]::Round($_.FreeSpace/1GB,1)
    $used = [math]::Round($size - $free,1)

    "$($_.DeviceID): Used $used GB / $size GB"
}) -join " || "

# SSD / HDD / NVMe
$physical = Get-PhysicalDisk -ErrorAction SilentlyContinue
$diskType = if ($physical) { ($physical.MediaType -join " | ") } else { "Unknown" }

# ================= NETWORK =================
$net = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object {$_.IPEnabled}

# ================= MONITOR =================
$monitors = Get-CimInstance -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue

$monitorInfo = ($monitors | ForEach-Object {
    $name = ($_.UserFriendlyName | ForEach-Object {[char]$_}) -join ""
    $man = ($_.ManufacturerName | ForEach-Object {[char]$_}) -join ""
    "$man $name"
}) -join " | "

# ================= AUDIO (HEADSET / SPEAKERS) =================
$audio = Get-CimInstance Win32_SoundDevice

# ================= CAMERA =================
$camera = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.Name -match "Camera|Webcam"
}

$cameraList = ($camera.Name -join " | ")

# ================= KEYBOARD / MOUSE =================
$keyboard = Get-CimInstance Win32_Keyboard
$mouse = Get-CimInstance Win32_PointingDevice

$keyboardList = ($keyboard.Name -join " | ")
$mouseList = ($mouse.Name -join " | ")

# ================= DOCKING STATION (BEST EFFORT) =================
$dock = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.Name -match "Dock|USB-C|Thunderbolt"
}

$dockList = ($dock.Name -join " | ")

# ================= INSTALLED APPS =================
$apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
                        HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* `
    -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName }

$appList = ($apps.DisplayName -join " | ")

# ================= OFFICE =================
$office = Get-CimInstance SoftwareLicensingProduct |
    Where-Object { $_.Name -like "*Office*" -and $_.PartialProductKey }

$officeStatus = if ($office) { "Installed & Activated" } else { "Not Installed / Unknown" }

# ================= FINAL OUTPUT =================
$data = [PSCustomObject]@{

    # EMPLOYEE
    EmployeeName = $employeeName
    PhoneNumber = if ($phone -eq "") { "Not Provided" } else { $phone }
    Email = $email

    # DEVICE TYPE
    DeviceType = $deviceType

    # SYSTEM
    ComputerName = $env:COMPUTERNAME
    User = $env:USERNAME

    Manufacturer = $cs.Manufacturer
    Model = $cs.Model
    SerialNumber = $bios.SerialNumber

    OS = $os.Caption

    # CPU / GPU
    CPU = $cpu.Name
    GPU = ($gpu.Name -join "; ")

    # RAM
    RAM_Total_GB = $ramTotal
    RAM_Details = $ramDetails

    # STORAGE
    Disk_Details = $diskDetails
    Disk_Type = $diskType
    Drive_Usage = $driveUsage

    # NETWORK
    IP = ($net.IPAddress -join "; ")
    MAC = ($net.MACAddress -join "; ")

    # PERIPHERALS
    Keyboard = $keyboardList
    Mouse = $mouseList
    Camera = $cameraList
    Audio = ($audio.Name -join "; ")
    Monitor = $monitorInfo
    DockStation = $dockList

    # SOFTWARE
    InstalledAppsCount = $apps.Count
    InstalledApps = $appList

    # OFFICE
    OfficeStatus = $officeStatus
}

$data | Export-Csv -Path $output -NoTypeInformation -Encoding UTF8

Write-Host "SUCCESS: FULL IT inventory saved to $output"

$root = $PSScriptRoot
if (-not $root) { $root = Get-Location }

$outputFile = Join-Path $root "merged.csv"
$logFile = Join-Path $root "processed.log"

# load already processed files
$processed = @()
if (Test-Path $logFile) {
    $processed = Get-Content $logFile
}

# get all CSVs except merged file
$files = Get-ChildItem -Path $root -Recurse -Filter *.csv |
    Where-Object { $_.FullName -ne $outputFile }

# filter ONLY new files
$newFiles = $files | Where-Object { $processed -notcontains $_.FullName }

foreach ($file in $newFiles) {

    $data = Import-Csv $file.FullName |
        Select-Object *, @{Name="SourceFile";Expression={$file.FullName}}

    # append safely
    if (-not (Test-Path $outputFile)) {
        $data | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8
    }
    else {
        $data | Export-Csv $outputFile -NoTypeInformation -Append -Encoding UTF8
    }

    # mark file as processed
    Add-Content $logFile $file.FullName
}

