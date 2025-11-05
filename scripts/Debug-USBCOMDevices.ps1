#Requires -Version 5.0
<#
.SYNOPSIS
    Diagnostic tool for USB COM port device troubleshooting.

.DESCRIPTION
    Gathers comprehensive information about USB devices, COM ports, and device IDs
    to help debug the COM port assignment script. Does not make any changes.

.PARAMETER VendorIDSearch
    Optional: Specific Vendor ID to search for (e.g., "0403")

.PARAMETER ProductIDSearch
    Optional: Specific Product ID to search for (e.g., "6001")

.PARAMETER OutputPath
    Optional: Path to save diagnostic report (default: C:\Temp)

.NOTES
    File Name      : Debug-USBCOMDevices.ps1
    Author         : VDI Admin
    Prerequisite   : PowerShell 5.0
    Version        : 1.0
    Location       : C:\Imaging\Scripts\
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$VendorIDSearch,
    
    [Parameter(Mandatory=$false)]
    [string]$ProductIDSearch,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "C:\Temp"
)

# Set output encoding to UTF-8
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # Silently continue if no console is available
}
$OutputEncoding = [System.Text.Encoding]::UTF8

# Configuration
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFile = Join-Path -Path $OutputPath -ChildPath "USB-COM-Diagnostic-Report_$timestamp.txt"

function Write-DiagnosticOutput {
    param(
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewLine
    )
    
    # Write to console
    if ($NoNewLine) {
        Write-Host $Message -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $Color
    }
    
    # Write to file
    Add-Content -Path $reportFile -Value $Message -Encoding UTF8
}

function Write-SectionHeader {
    param([string]$Title)
    
    Write-DiagnosticOutput "`n"
    Write-DiagnosticOutput "========================================" -Color Cyan
    Write-DiagnosticOutput $Title -Color Cyan
    Write-DiagnosticOutput "========================================" -Color Cyan
}

# Start diagnostic
Clear-Host
Write-SectionHeader "USB COM Port Diagnostic Tool"
Write-DiagnosticOutput "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-DiagnosticOutput "Computer: $env:COMPUTERNAME"
Write-DiagnosticOutput "User: $env:USERNAME"
Write-DiagnosticOutput "Report: $reportFile"

# System Information
Write-SectionHeader "SYSTEM INFORMATION"

Write-DiagnosticOutput "`nOperating System:"
$os = Get-CimInstance Win32_OperatingSystem
Write-DiagnosticOutput "  OS Name: $($os.Caption)"
Write-DiagnosticOutput "  Version: $($os.Version)"
Write-DiagnosticOutput "  Architecture: $($os.OSArchitecture)"

Write-DiagnosticOutput "`nPowerShell Version:"
Write-DiagnosticOutput "  PSVersion: $($PSVersionTable.PSVersion)"
Write-DiagnosticOutput "  PSEdition: $($PSVersionTable.PSEdition)"

Write-DiagnosticOutput "`nExecution Context:"
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Write-DiagnosticOutput "  Running as Admin: $isAdmin"
Write-DiagnosticOutput "  User Context: $env:USERNAME"

# Check for Horizon Agent
Write-SectionHeader "OMNISSA HORIZON AGENT"

$horizonAgentPath = "HKLM:\SOFTWARE\VMware, Inc.\VMware VDM"
if (Test-Path $horizonAgentPath) {
    Write-DiagnosticOutput "`nHorizon Agent: Installed" -Color Green
    
    $agentVersion = Get-ItemProperty -Path "$horizonAgentPath\Agent" -Name "Version" -ErrorAction SilentlyContinue
    if ($agentVersion) {
        Write-DiagnosticOutput "  Version: $($agentVersion.Version)"
    }
    
    # Check USB settings
    $usbPath = "$horizonAgentPath\Agent\USB"
    if (Test-Path $usbPath) {
        Write-DiagnosticOutput "`nUSB Configuration:"
        Get-ItemProperty -Path $usbPath | 
            Get-Member -MemberType NoteProperty | 
            Where-Object {$_.Name -notlike "PS*"} | 
            ForEach-Object {
                $value = (Get-ItemProperty -Path $usbPath).$($_.Name)
                Write-DiagnosticOutput "  $($_.Name): $value"
            }
    }
    
    # Check USB Arbitrator Service
    $usbService = Get-Service -Name "VMwareUSBArbitrationService" -ErrorAction SilentlyContinue
    if ($usbService) {
        Write-DiagnosticOutput "`nUSB Arbitrator Service:"
        Write-DiagnosticOutput "  Status: $($usbService.Status)" -Color $(if ($usbService.Status -eq 'Running') {'Green'} else {'Red'})
        Write-DiagnosticOutput "  StartType: $($usbService.StartType)"
    }
} else {
    Write-DiagnosticOutput "`nHorizon Agent: Not Detected" -Color Yellow
}

# All USB Devices
Write-SectionHeader "ALL USB DEVICES"

Write-DiagnosticOutput "`nScanning for USB devices..."
$allUSBDevices = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.DeviceID -match "^USB"
}

Write-DiagnosticOutput "Found $($allUSBDevices.Count) USB devices"

foreach ($device in $allUSBDevices | Sort-Object Name) {
    $deviceID = $device.DeviceID
    $color = "White"
    
    # Extract VendorID and ProductID
    if ($deviceID -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
        $foundVendorID = $matches[1]
        $foundProductID = $matches[2]
        
        # Highlight if matches search criteria
        if (($VendorIDSearch -and $foundVendorID -eq $VendorIDSearch) -or 
            ($ProductIDSearch -and $foundProductID -eq $ProductIDSearch)) {
            $color = "Yellow"
        }
        
        Write-DiagnosticOutput "`n  Name: $($device.Name)" -Color $color
        Write-DiagnosticOutput "  VendorID: $foundVendorID" -Color $color
        Write-DiagnosticOutput "  ProductID: $foundProductID" -Color $color
        Write-DiagnosticOutput "  DeviceID: $deviceID"
        Write-DiagnosticOutput "  Status: $($device.Status)"
        
        if ($device.ConfigManagerErrorCode -ne 0) {
            Write-DiagnosticOutput "  Error Code: $($device.ConfigManagerErrorCode)" -Color Red
        }
    }
}

# COM Port Devices
Write-SectionHeader "COM PORT DEVICES"

Write-DiagnosticOutput "`nScanning for COM port devices..."
$comDevices = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.Name -match 'COM\d+'
}

Write-DiagnosticOutput "Found $($comDevices.Count) COM port devices"

foreach ($device in $comDevices | Sort-Object Name) {
    $deviceID = $device.DeviceID
    $color = "White"
    
    Write-DiagnosticOutput "`n  Name: $($device.Name)"
    
    # Extract COM port number
    if ($device.Name -match '(COM\d+)') {
        $comPort = $matches[1]
        Write-DiagnosticOutput "  COM Port: $comPort" -Color Cyan
    }
    
    # Extract VendorID and ProductID if USB device
    if ($deviceID -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
        $foundVendorID = $matches[1]
        $foundProductID = $matches[2]
        
        # Highlight if matches search criteria
        if (($VendorIDSearch -and $foundVendorID -eq $VendorIDSearch) -or 
            ($ProductIDSearch -and $foundProductID -eq $ProductIDSearch)) {
            $color = "Yellow"
        }
        
        Write-DiagnosticOutput "  VendorID: $foundVendorID" -Color $color
        Write-DiagnosticOutput "  ProductID: $foundProductID" -Color $color
    }
    
    Write-DiagnosticOutput "  DeviceID: $deviceID"
    Write-DiagnosticOutput "  Status: $($device.Status)"
    
    if ($device.ConfigManagerErrorCode -ne 0) {
        Write-DiagnosticOutput "  Error Code: $($device.ConfigManagerErrorCode)" -Color Red
    }
}

# Target Devices (if specified)
if ($VendorIDSearch -or $ProductIDSearch) {
    Write-SectionHeader "CHECKING TARGET DEVICES"
    
    $searchCriteria = @()
    if ($VendorIDSearch) { $searchCriteria += "VendorID: $VendorIDSearch" }
    if ($ProductIDSearch) { $searchCriteria += "ProductID: $ProductIDSearch" }
    
    Write-DiagnosticOutput "`nSearching for: $($searchCriteria -join ', ')"
    
    $targetDevices = Get-CimInstance Win32_PnPEntity | Where-Object {
        $deviceID = $_.DeviceID
        
        if ($deviceID -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
            $foundVendorID = $matches[1]
            $foundProductID = $matches[2]
            
            $vendorIDMatch = -not $VendorIDSearch -or ($foundVendorID -eq $VendorIDSearch)
            $productIDMatch = -not $ProductIDSearch -or ($foundProductID -eq $ProductIDSearch)
            
            return ($vendorIDMatch -and $productIDMatch)
        }
        
        return $false
    }
    
    if ($targetDevices) {
        Write-DiagnosticOutput "Found $($targetDevices.Count) matching device(s)" -Color Green
        
        foreach ($device in $targetDevices) {
            Write-DiagnosticOutput "`n  Name: $($device.Name)" -Color Green
            Write-DiagnosticOutput "  DeviceID: $($device.DeviceID)"
            Write-DiagnosticOutput "  Status: $($device.Status)"
            
            # Check if it has a COM port
            if ($device.Name -match '(COM\d+)') {
                Write-DiagnosticOutput "  COM Port: $($matches[1])" -Color Green
            } else {
                Write-DiagnosticOutput "  COM Port: Not assigned" -Color Yellow
            }
            
            if ($device.ConfigManagerErrorCode -ne 0) {
                Write-DiagnosticOutput "  Error Code: $($device.ConfigManagerErrorCode)" -Color Red
            }
        }
    } else {
        Write-DiagnosticOutput "No matching devices found" -Color Red
    }
}

# Registry COM Port Assignments
Write-SectionHeader "REGISTRY COM PORT ASSIGNMENTS"

Write-DiagnosticOutput "`nScanning registry for COM port assignments..."

$enumPath = "HKLM:\SYSTEM\CurrentControlSet\Enum"
$foundPorts = @()

Get-ChildItem -Path "$enumPath\USB" -Recurse -ErrorAction SilentlyContinue | 
    Where-Object {$_.PSChildName -eq "Device Parameters"} | 
    ForEach-Object {
        try {
            $portName = (Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue).PortName
            if ($portName) {
                $parentPath = Split-Path -Path $_.PSPath -Parent
                $deviceIDParts = $parentPath -replace 'Microsoft.PowerShell.Core\\Registry::HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\',''
                
                $foundPorts += [PSCustomObject]@{
                    COMPort = $portName
                    DevicePath = $deviceIDParts
                }
            }
        } catch {
            # Continue on error
        }
    }

Write-DiagnosticOutput "Found $($foundPorts.Count) COM port assignments in registry"

foreach ($port in $foundPorts | Sort-Object COMPort) {
    Write-DiagnosticOutput "`n  COM Port: $($port.COMPort)" -Color Cyan
    Write-DiagnosticOutput "  Device Path: $($port.DevicePath)"
    
    # Extract VendorID and ProductID
    if ($port.DevicePath -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
        Write-DiagnosticOutput "  VendorID: $($matches[1])"
        Write-DiagnosticOutput "  ProductID: $($matches[2])"
    }
}

# Check for Device Problems
Write-SectionHeader "PROBLEM DEVICES"

Write-DiagnosticOutput "`nChecking for devices with errors..."

$problemDevices = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.ConfigManagerErrorCode -ne 0 -and $_.DeviceID -match "^USB"
}

if ($problemDevices) {
    Write-DiagnosticOutput "Found $($problemDevices.Count) device(s) with errors" -Color Red
    
    foreach ($device in $problemDevices) {
        Write-DiagnosticOutput "`n  Name: $($device.Name)" -Color Red
        Write-DiagnosticOutput "  DeviceID: $($device.DeviceID)"
        Write-DiagnosticOutput "  Error Code: $($device.ConfigManagerErrorCode)" -Color Red
        Write-DiagnosticOutput "  Status: $($device.Status)"
    }
} else {
    Write-DiagnosticOutput "No problem devices found" -Color Green
}

# Check Script Logs
Write-SectionHeader "USB COM PORT ASSIGNMENT SCRIPT"

Write-DiagnosticOutput "`nChecking for script execution..."

$logPath = "C:\Temp"
$logFiles = Get-ChildItem -Path $logPath -Filter "USBCOMPortAssignment_*.log" -ErrorAction SilentlyContinue | 
    Sort-Object LastWriteTime -Descending

if ($logFiles) {
    Write-DiagnosticOutput "Found $($logFiles.Count) log file(s)" -Color Green
    
    $latestLog = $logFiles[0]
    Write-DiagnosticOutput "`nLatest log file:"
    Write-DiagnosticOutput "  File: $($latestLog.Name)"
    Write-DiagnosticOutput "  Date: $($latestLog.LastWriteTime)"
    Write-DiagnosticOutput "  Size: $($latestLog.Length) bytes"
    
    Write-DiagnosticOutput "`nRecent log entries (last 20 lines):"
    $logContent = Get-Content $latestLog.FullName -Tail 20
    foreach ($line in $logContent) {
        $color = "White"
        if ($line -match "\[ERROR\]") { $color = "Red" }
        elseif ($line -match "\[WARNING\]") { $color = "Yellow" }
        elseif ($line -match "\[SUCCESS\]") { $color = "Green" }
        
        Write-DiagnosticOutput "  $line" -Color $color
    }
    
    # Count errors and warnings
    $allContent = Get-Content $latestLog.FullName
    $errorCount = ($allContent | Select-String -Pattern "\[ERROR\]").Count
    $warningCount = ($allContent | Select-String -Pattern "\[WARNING\]").Count
    $successCount = ($allContent | Select-String -Pattern "\[SUCCESS\]").Count
    
    Write-DiagnosticOutput "`nLog Summary:"
    Write-DiagnosticOutput "  Errors: $errorCount" -Color $(if ($errorCount -gt 0) {'Red'} else {'Green'})
    Write-DiagnosticOutput "  Warnings: $warningCount" -Color $(if ($warningCount -gt 0) {'Yellow'} else {'Green'})
    Write-DiagnosticOutput "  Successes: $successCount" -Color Green
} else {
    Write-DiagnosticOutput "No log files found" -Color Yellow
    Write-DiagnosticOutput "Script may not have run yet or logs may have been cleared"
}

# Check Scheduled Task
Write-DiagnosticOutput "`nChecking scheduled task..."

$taskName = "USB-COM-Port-Assignment"
$task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($task) {
    Write-DiagnosticOutput "Scheduled task found: $taskName" -Color Green
    Write-DiagnosticOutput "  State: $($task.State)"
    
    $taskInfo = Get-ScheduledTaskInfo -TaskName $taskName
    Write-DiagnosticOutput "  Last Run Time: $($taskInfo.LastRunTime)"
    Write-DiagnosticOutput "  Last Result: $($taskInfo.LastTaskResult) $(if ($taskInfo.LastTaskResult -eq 0) {'(Success)'} else {'(Failed)'})"
    Write-DiagnosticOutput "  Next Run Time: $($taskInfo.NextRunTime)"
} else {
    Write-DiagnosticOutput "Scheduled task not found: $taskName" -Color Red
}

# Test Script Logic
Write-SectionHeader "TESTING SCRIPT LOGIC"

Write-DiagnosticOutput "`nSimulating script device detection logic..."

# Target devices from script
$targetDevices = @(
    @{Name = "Topaz T-LBK462-BSB-RC"; VendorID = "0403"; ProductID = "6001"; TargetPort = "COM5"},
    @{Name = "Ingenico LANE3000"; VendorID = "0B00"; ProductID = "0084"; TargetPort = "COM21"}
)

foreach ($target in $targetDevices) {
    Write-DiagnosticOutput "`nChecking: $($target.Name)"
    Write-DiagnosticOutput "  Target VendorID: $($target.VendorID)"
    Write-DiagnosticOutput "  Target ProductID: $($target.ProductID)"
    Write-DiagnosticOutput "  Target COM Port: $($target.TargetPort)"
    
    # Try to find device
    $found = Get-CimInstance Win32_PnPEntity | Where-Object {
        $deviceID = $_.DeviceID
        
        if ($deviceID -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
            $foundVendorID = $matches[1]
            $foundProductID = $matches[2]
            
            return (($foundVendorID -eq $target.VendorID) -and ($foundProductID -eq $target.ProductID))
        }
        
        return $false
    }
    
    if ($found) {
        Write-DiagnosticOutput "  Device FOUND" -Color Green
        Write-DiagnosticOutput "  Current Name: $($found.Name)"
        
        # Check for COM port
        if ($found.Name -match '(COM\d+)') {
            $currentPort = $matches[1]
            Write-DiagnosticOutput "  Current COM Port: $currentPort"
            
            if ($currentPort -eq $target.TargetPort) {
                Write-DiagnosticOutput "  Status: On correct port" -Color Green
            } else {
                Write-DiagnosticOutput "  Status: On wrong port - needs reassignment" -Color Yellow
            }
        } else {
            Write-DiagnosticOutput "  Current COM Port: Not assigned" -Color Yellow
            Write-DiagnosticOutput "  Status: No COM port - check device driver" -Color Red
        }
    } else {
        Write-DiagnosticOutput "  Device NOT FOUND" -Color Red
        Write-DiagnosticOutput "  Possible reasons:"
        Write-DiagnosticOutput "    - Device not connected"
        Write-DiagnosticOutput "    - USB redirection not working"
        Write-DiagnosticOutput "    - Device not whitelisted in Horizon"
        Write-DiagnosticOutput "    - Wrong VendorID/ProductID in script"
    }
}

# Summary and Recommendations
Write-SectionHeader "SUMMARY AND RECOMMENDATIONS"

Write-DiagnosticOutput "`nDiagnostic Summary:"
Write-DiagnosticOutput "  Total USB Devices: $($allUSBDevices.Count)"
Write-DiagnosticOutput "  COM Port Devices: $($comDevices.Count)"
Write-DiagnosticOutput "  Problem Devices: $($problemDevices.Count)"

Write-DiagnosticOutput "`nRecommendations:"

if (-not $isAdmin) {
    Write-DiagnosticOutput "  - Script requires elevated privileges (SYSTEM context)" -Color Yellow
}

if ($problemDevices.Count -gt 0) {
    Write-DiagnosticOutput "  - Investigate and resolve $($problemDevices.Count) problem device(s)" -Color Red
}

if (-not (Test-Path $horizonAgentPath)) {
    Write-DiagnosticOutput "  - Horizon Agent not detected - verify VDI environment" -Color Yellow
}

if (-not $task) {
    Write-DiagnosticOutput "  - Scheduled task not found - script may not run automatically" -Color Yellow
}

if ($logFiles.Count -eq 0) {
    Write-DiagnosticOutput "  - No log files found - script has not run or logs were cleared" -Color Yellow
}

Write-DiagnosticOutput "`n"
Write-DiagnosticOutput "Diagnostic complete! Report saved to:" -Color Green
Write-DiagnosticOutput $reportFile -Color Cyan
Write-DiagnosticOutput "`n"
