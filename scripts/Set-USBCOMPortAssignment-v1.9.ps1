#Requires -Version 5.0
<#
.SYNOPSIS
    Manages static COM port assignments for USB devices in Omnissa Horizon VDI.

.DESCRIPTION
    Monitors and assigns specific COM ports to USB devices based on VendorID/ProductID.
    Designed for non-persistent VDI with USB redirection.
    - Topaz T-LBK462-BSB-RC (VendorID:0403, ProductID:6001) -> COM5
    - Ingenico LANE3000 (VendorID:0B00, ProductID:0084) -> COM21
    Also updates the device friendly name in Device Manager to reflect the new COM port.

.NOTES
    File Name      : Set-USBCOMPortAssignment.ps1
    Author         : VDI Admin
    Prerequisite   : PowerShell 5.0, Elevated permissions (SYSTEM account via scheduled task)
    Version        : 1.9
    Location       : C:\Imaging\Scripts\
    Fix v1.2       : Changed from WMI -Filter to Where-Object for reliability
    Fix v1.3       : Enhanced registry path detection with fallback search
    Fix v1.4       : Renamed VID/PID to VendorID/ProductID to avoid variable conflicts
    Fix v1.5       : Added friendly name update in Device Manager
    Fix v1.6       : Added configurable COM port settings (baud rate, parity, flow control, etc.)
    Fix v1.7       : Updated with manufacturer-specified COM settings for both devices
    Fix v1.8       : Fixed registry path resolution in Set-COMPortSettings and Set-COMPortAssignment
    Fix v1.9       : CRITICAL - Dual registry writes for Device Manager visibility + timing fixes
    Note           : Requires SYSTEM context for registry write access to device parameters
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$MaxWaitSeconds = 60,
    
    [Parameter(Mandatory=$false)]
    [int]$CheckIntervalSeconds = 2,
    
    [Parameter(Mandatory=$false)]
    [switch]$ContinuousMonitoring
)

# Set output encoding to UTF-8 (with error handling for console-less environments)
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # Silently continue if no console is available
}
$OutputEncoding = [System.Text.Encoding]::UTF8

# Configuration
$LogPath = "C:\Temp"
$LogFile = Join-Path -Path $LogPath -ChildPath "USBCOMPortAssignment_$(Get-Date -Format 'yyyyMMdd').log"

# Device configuration: VendorID, ProductID, Target COM Port, COM Settings
# 
# COM Settings Format: "BaudRate,DataBits,Parity,StopBits,FlowControl,UseFIFO,RxBuffer,TxBuffer"
#   - BaudRate: 110, 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200, 128000, 256000
#   - DataBits: 5, 6, 7, 8
#   - Parity: 0=None, 1=Odd, 2=Even, 3=Mark, 4=Space
#   - StopBits: 0=1bit, 1=1.5bits, 2=2bits
#   - FlowControl: 0=None, 1=Xon/Xoff, 2=Hardware, 3=Both
#   - UseFIFO: 0=Disabled, 1=Enabled
#   - RxBuffer: 1-14 (Registry value, actual buffer size varies)
#   - TxBuffer: 1-14 (Registry value, actual buffer size varies)
#
# Use "default" to skip custom COM port settings
#
$DeviceConfig = @(
    @{
        Name = "Topaz T-LBK462-BSB-RC"
        VendorID = "0403"
        ProductID = "6001"
        TargetCOMPort = "COM5"
        # Topaz BSB signature pad specifications:
        # 19200 baud, 8 data bits, Odd parity, 1 stop bit, No flow control, FIFO enabled, max buffers
        COMSettings = "19200,8,1,0,0,1,14,14"
    },
    @{
        Name = "Ingenico LANE3000"
        VendorID = "0B00"
        ProductID = "0084"
        TargetCOMPort = "COM21"
        # Ingenico LANE3000 specifications:
        # 115200 baud, 8 data bits, No parity, 1 stop bit, No flow control, FIFO enabled, max buffers
        COMSettings = "115200,8,0,0,0,1,14,14"
    }
)

#region Functions

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Ensure log directory exists
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logMessage -Encoding UTF8
    
    # Write to console with color
    $color = switch ($Level) {
        'INFO'    { 'White' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        default   { 'White' }
    }
    
    try {
        Write-Host $logMessage -ForegroundColor $color
    } catch {
        # Silently continue if no console is available
    }
}

function Get-USBDeviceCOMPort {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetVendorID,
        
        [Parameter(Mandatory=$true)]
        [string]$TargetProductID
    )
    
    try {
        # Get all COM port devices
        $comDevices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
            $_.Name -match 'COM\d+'
        }
        
        foreach ($device in $comDevices) {
            $deviceID = $device.DeviceID
            
            # Extract VendorID and ProductID from DeviceID
            if ($deviceID -match "VID_([0-9A-F]{4}).*PID_([0-9A-F]{4})") {
                $foundVendorID = $matches[1]
                $foundProductID = $matches[2]
                
                # Compare with target (case-insensitive)
                if (($foundVendorID -eq $TargetVendorID) -and ($foundProductID -eq $TargetProductID)) {
                    # Extract COM port number from device name
                    if ($device.Name -match '(COM\d+)') {
                        $comPort = $matches[1]
                        
                        return @{
                            DeviceID = $deviceID
                            Name = $device.Name
                            COMPort = $comPort
                            Status = $device.Status
                            ConfigManagerErrorCode = $device.ConfigManagerErrorCode
                        }
                    }
                }
            }
        }
        
        return $null
    }
    catch {
        Write-Log -Message "Error getting USB device COM port: $_" -Level ERROR
        return $null
    }
}

function Get-COMPortInUse {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$COMPort
    )
    
    try {
        # Check if COM port is in use
        $comDevices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
            $_.Name -match $COMPort
        }
        
        return ($null -ne $comDevices)
    }
    catch {
        Write-Log -Message "Error checking COM port usage: $_" -Level ERROR
        return $false
    }
}

function Set-COMPortSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DeviceID,
        
        [Parameter(Mandatory=$true)]
        [string]$COMPort,
        
        [Parameter(Mandatory=$true)]
        [string]$SettingsString
    )
    
    # Skip if default settings
    if ($SettingsString -eq "default") {
        Write-Log -Message "Using default COM port settings for $COMPort" -Level INFO
        return $true
    }
    
    try {
        # Parse the settings string
        $settings = $SettingsString -split ','
        
        if ($settings.Count -ne 8) {
            Write-Log -Message "Invalid COM settings format. Expected 8 values, got $($settings.Count)" -Level ERROR
            return $false
        }
        
        $baudRate = [int]$settings[0]
        $dataBits = [int]$settings[1]
        $parity = [int]$settings[2]
        $stopBits = [int]$settings[3]
        $flowControl = [int]$settings[4]
        $useFIFO = [int]$settings[5]
        $rxBuffer = [int]$settings[6]
        $txBuffer = [int]$settings[7]
        
        # Find the device parameters registry path
        # DeviceID format: USB\VID_0403&PID_6001\1234567890
        $enumPath = "HKLM:\SYSTEM\CurrentControlSet\Enum"
        $devicePath = $null
        
        # Construct path directly from DeviceID
        $deviceIdParts = $DeviceID -split '\\'
        if ($deviceIdParts.Count -ge 2) {
            # Build the full registry path
            $constructedPath = Join-Path -Path $enumPath -ChildPath ($deviceIdParts -join '\')
            if (Test-Path -Path $constructedPath) {
                $devicePath = $constructedPath
                Write-Log -Message "Found device registry path: $devicePath" -Level INFO
            }
        }
        
        if (-not $devicePath) {
            # Fallback: Search through USB devices
            $usbPath = Join-Path -Path $enumPath -ChildPath "USB"
            if (Test-Path -Path $usbPath) {
                Get-ChildItem -Path $usbPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                    # Compare the registry path to the DeviceID
                    $registryPath = $_.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\', ''
                    $registryPath = $registryPath -replace '\\', '\'
                    
                    if ($registryPath -eq $DeviceID) {
                        $devicePath = $_.PSPath
                        Write-Log -Message "Found device via search: $devicePath" -Level INFO
                        return
                    }
                }
            }
        }
        
        if (-not $devicePath) {
            Write-Log -Message "Could not find registry path for device: $DeviceID" -Level WARNING
            return $false
        }
        
        # Get Device Parameters path
        $deviceParamsPath = Join-Path -Path $devicePath -ChildPath "Device Parameters"
        
        if (-not (Test-Path -Path $deviceParamsPath)) {
            Write-Log -Message "Device Parameters path does not exist: $deviceParamsPath" -Level WARNING
            return $false
        }
        
        Write-Log -Message "Applying COM settings to $COMPort" -Level INFO
        Write-Log -Message "  BaudRate: $baudRate, DataBits: $dataBits, Parity: $parity, StopBits: $stopBits" -Level INFO
        Write-Log -Message "  FlowControl: $flowControl, FIFO: $useFIFO, RxBuffer: $rxBuffer, TxBuffer: $txBuffer" -Level INFO
        
        # CRITICAL: Write to BOTH registry locations for persistence and Device Manager visibility
        
        # Location 1: Device Parameters (device-specific settings)
        Set-ItemProperty -Path $deviceParamsPath -Name "BaudRate" -Value $baudRate -Type DWord -Force
        Set-ItemProperty -Path $deviceParamsPath -Name "DataBits" -Value $dataBits -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "Parity" -Value $parity -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "StopBits" -Value $stopBits -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "FlowControl" -Value $flowControl -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "FifoEnable" -Value $useFIFO -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "RxFIFO" -Value $rxBuffer -Type DWord -Force -ErrorAction SilentlyContinue
        Set-ItemProperty -Path $deviceParamsPath -Name "TxFIFO" -Value $txBuffer -Type DWord -Force -ErrorAction SilentlyContinue
        
        # Location 2: Windows Ports registry (what Device Manager displays)
        # Format: "BaudRate,Parity,DataBits,StopBits"
        # Parity conversion: 0=n, 1=o, 2=e, 3=m, 4=s
        # StopBits conversion: 0=1, 1=1.5, 2=2
        $portsPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports"
        $parityChar = switch ($parity) {
            0 { 'n' }
            1 { 'o' }
            2 { 'e' }
            3 { 'm' }
            4 { 's' }
            default { 'n' }
        }
        $stopBitsValue = switch ($stopBits) {
            0 { '1' }
            1 { '1.5' }
            2 { '2' }
            default { '1' }
        }
        $portsValue = "$baudRate,$parityChar,$dataBits,$stopBitsValue"
        
        # Ensure Ports key exists
        if (-not (Test-Path -Path $portsPath)) {
            New-Item -Path $portsPath -Force | Out-Null
        }
        
        # Set the Ports registry value (note the colon after COM port name)
        Set-ItemProperty -Path $portsPath -Name "${COMPort}:" -Value $portsValue -Type String -Force
        Write-Log -Message "Set Windows Ports registry: ${COMPort}: = $portsValue" -Level INFO
        
        Write-Log -Message "Successfully applied COM port settings for $COMPort" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log -Message "Error setting COM port settings: $_" -Level ERROR
        return $false
    }
}

function Set-COMPortAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DeviceID,
        
        [Parameter(Mandatory=$true)]
        [string]$CurrentCOMPort,
        
        [Parameter(Mandatory=$true)]
        [string]$TargetCOMPort,
        
        [Parameter(Mandatory=$true)]
        [string]$DeviceName
    )
    
    try {
        Write-Log -Message "Attempting to reassign $CurrentCOMPort to $TargetCOMPort for device: $DeviceName" -Level INFO
        
        # Find the device parameters registry path
        # DeviceID format: USB\VID_0403&PID_6001\1234567890
        $enumPath = "HKLM:\SYSTEM\CurrentControlSet\Enum"
        $devicePath = $null
        
        # Construct path directly from DeviceID
        $deviceIdParts = $DeviceID -split '\\'
        if ($deviceIdParts.Count -ge 2) {
            # Build the full registry path
            $constructedPath = Join-Path -Path $enumPath -ChildPath ($deviceIdParts -join '\')
            if (Test-Path -Path $constructedPath) {
                $devicePath = $constructedPath
                Write-Log -Message "Found device registry path: $devicePath" -Level INFO
            }
        }
        
        if (-not $devicePath) {
            # Fallback: Search through USB devices
            $usbPath = Join-Path -Path $enumPath -ChildPath "USB"
            if (Test-Path -Path $usbPath) {
                Get-ChildItem -Path $usbPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                    # Compare the registry path to the DeviceID
                    $registryPath = $_.PSPath -replace 'Microsoft.PowerShell.Core\\Registry::HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Enum\\', ''
                    $registryPath = $registryPath -replace '\\', '\'
                    
                    if ($registryPath -eq $DeviceID) {
                        $devicePath = $_.PSPath
                        Write-Log -Message "Found device via search: $devicePath" -Level INFO
                        return
                    }
                }
            }
        }
        
        if (-not $devicePath) {
            Write-Log -Message "Could not find registry path for device: $DeviceID" -Level ERROR
            return $false
        }
        
        # Get Device Parameters path
        $deviceParamsPath = Join-Path -Path $devicePath -ChildPath "Device Parameters"
        
        if (-not (Test-Path -Path $deviceParamsPath)) {
            Write-Log -Message "Device Parameters path does not exist: $deviceParamsPath" -Level ERROR
            return $false
        }
        
        # Set the new COM port name
        Set-ItemProperty -Path $deviceParamsPath -Name "PortName" -Value $TargetCOMPort -Force
        Write-Log -Message "Successfully updated registry: $CurrentCOMPort -> $TargetCOMPort" -Level SUCCESS
        
        # Update the friendly name in Device Manager
        $friendlyNamePath = $devicePath
        $newFriendlyName = "$DeviceName ($TargetCOMPort)"
        
        try {
            Set-ItemProperty -Path $friendlyNamePath -Name "FriendlyName" -Value $newFriendlyName -Force
            Write-Log -Message "Successfully updated friendly name to: $newFriendlyName" -Level SUCCESS
        }
        catch {
            Write-Log -Message "Warning: Could not update friendly name: $_" -Level WARNING
        }
        
        return $true
    }
    catch {
        Write-Log -Message "Error setting COM port assignment: $_" -Level ERROR
        return $false
    }
}

function Restart-PnPDevice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$DeviceID
    )
    
    try {
        Write-Log -Message "Restarting device: $DeviceID" -Level INFO
        
        # Get the device
        $device = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
            $_.DeviceID -eq $DeviceID
        }
        
        if ($null -eq $device) {
            Write-Log -Message "Device not found for restart: $DeviceID" -Level WARNING
            return $false
        }
        
        # Disable the device
        $device | Invoke-CimMethod -MethodName Disable -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 2
        
        # Enable the device
        $device | Invoke-CimMethod -MethodName Enable -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 3
        
        Write-Log -Message "Successfully restarted device" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log -Message "Error restarting device: $_" -Level ERROR
        return $false
    }
}

function Clear-COMPortConflict {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetCOMPort,
        
        [Parameter(Mandatory=$true)]
        [string]$ExcludeDeviceID
    )
    
    try {
        Write-Log -Message "Checking for conflicts on $TargetCOMPort" -Level INFO
        
        # Find any device currently using the target COM port
        $conflictingDevice = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object {
            ($_.Name -match $TargetCOMPort) -and ($_.DeviceID -ne $ExcludeDeviceID)
        }
        
        if ($null -ne $conflictingDevice) {
            Write-Log -Message "Found conflicting device on $TargetCOMPort : $($conflictingDevice.Name)" -Level WARNING
            Write-Log -Message "Attempting to clear conflict..." -Level INFO
            
            # Try to reassign the conflicting device to a different COM port
            # Find the next available COM port (COM100 and above to avoid conflicts)
            $newPort = "COM100"
            $portNumber = 100
            
            while (Get-COMPortInUse -COMPort $newPort) {
                $portNumber++
                $newPort = "COM$portNumber"
                
                if ($portNumber -gt 200) {
                    Write-Log -Message "Could not find available COM port for conflicting device" -Level ERROR
                    return $false
                }
            }
            
            # Reassign the conflicting device
            if (Set-COMPortAssignment -DeviceID $conflictingDevice.DeviceID -CurrentCOMPort $TargetCOMPort -TargetCOMPort $newPort -DeviceName $conflictingDevice.Name) {
                Write-Log -Message "Successfully moved conflicting device to $newPort" -Level SUCCESS
                Restart-PnPDevice -DeviceID $conflictingDevice.DeviceID
                Start-Sleep -Seconds 2
                return $true
            }
            else {
                return $false
            }
        }
        else {
            Write-Log -Message "No conflicts found on $TargetCOMPort" -Level INFO
            return $true
        }
    }
    catch {
        Write-Log -Message "Error clearing COM port conflict: $_" -Level ERROR
        return $false
    }
}

function Process-DeviceAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )
    
    $deviceName = $Config.Name
    $vendorID = $Config.VendorID
    $productID = $Config.ProductID
    $targetPort = $Config.TargetCOMPort
    $comSettings = $Config.COMSettings
    
    Write-Log -Message "Processing device: $deviceName (VID:$vendorID, PID:$productID)" -Level INFO
    
    # Get current device info
    $deviceInfo = Get-USBDeviceCOMPort -TargetVendorID $vendorID -TargetProductID $productID
    
    if ($null -eq $deviceInfo) {
        Write-Log -Message "Device not found: $deviceName" -Level WARNING
        return $false
    }
    
    $currentPort = $deviceInfo.COMPort
    Write-Log -Message "Device found on $currentPort (Target: $targetPort)" -Level INFO
    
    # Check if device is already on the correct port
    if ($currentPort -eq $targetPort) {
        Write-Log -Message "Device already on correct port: $targetPort" -Level SUCCESS
        
        # Apply COM settings WITHOUT restart (port is already correct)
        if ($comSettings -ne "default") {
            Write-Log -Message "Applying COM settings without device restart..." -Level INFO
            Set-COMPortSettings -DeviceID $deviceInfo.DeviceID -COMPort $currentPort -SettingsString $comSettings
        }
        
        return $true
    }
    
    # Port needs to be changed
    Write-Log -Message "Port reassignment needed: $currentPort -> $targetPort" -Level INFO
    
    # Clear any conflicts on target port
    if (-not (Clear-COMPortConflict -TargetCOMPort $targetPort -ExcludeDeviceID $deviceInfo.DeviceID)) {
        Write-Log -Message "Failed to clear conflicts on $targetPort" -Level ERROR
        return $false
    }
    
    # Reassign the device to the target port
    if (Set-COMPortAssignment -DeviceID $deviceInfo.DeviceID -CurrentCOMPort $currentPort -TargetCOMPort $targetPort -DeviceName $deviceName) {
        Write-Log -Message "Successfully reassigned $deviceName from $currentPort to $targetPort" -Level SUCCESS
        
        # Restart the device to apply port change
        Write-Log -Message "Restarting device to apply port change..." -Level INFO
        Restart-PnPDevice -DeviceID $deviceInfo.DeviceID
        
        # Wait for driver to fully initialize (CRITICAL for COM settings persistence)
        Write-Log -Message "Waiting 5 seconds for driver initialization..." -Level INFO
        Start-Sleep -Seconds 5
        
        # NOW apply COM settings AFTER restart (so driver doesn't overwrite them)
        if ($comSettings -ne "default") {
            Write-Log -Message "Driver initialized - applying COM settings..." -Level INFO
            Set-COMPortSettings -DeviceID $deviceInfo.DeviceID -COMPort $targetPort -SettingsString $comSettings
        }
        
        # Verify the change
        $verifyInfo = Get-USBDeviceCOMPort -TargetVendorID $vendorID -TargetProductID $productID
        if ($null -ne $verifyInfo -and $verifyInfo.COMPort -eq $targetPort) {
            Write-Log -Message "Verified: $deviceName is now on $targetPort" -Level SUCCESS
            return $true
        }
        else {
            Write-Log -Message "Verification failed: Device may not have switched to $targetPort" -Level ERROR
            return $false
        }
    }
    else {
        Write-Log -Message "Failed to reassign $deviceName to $targetPort" -Level ERROR
        return $false
    }
}

#endregion

#region Main Script

Write-Log -Message "========== USB COM Port Assignment Script Started ==========" -Level INFO
Write-Log -Message "Script Version: 1.9" -Level INFO
Write-Log -Message "Running as: $env:USERNAME" -Level INFO
Write-Log -Message "Computer: $env:COMPUTERNAME" -Level INFO

# Check if running with elevated privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Log -Message "WARNING: Script is not running with elevated privileges. Registry writes may fail." -Level WARNING
}

# Wait for devices to be ready
Write-Log -Message "Waiting for USB devices to initialize..." -Level INFO
$elapsedSeconds = 0
$devicesReady = $false

while ($elapsedSeconds -lt $MaxWaitSeconds) {
    $foundCount = 0
    
    foreach ($config in $DeviceConfig) {
        $deviceInfo = Get-USBDeviceCOMPort -TargetVendorID $config.VendorID -TargetProductID $config.ProductID
        if ($null -ne $deviceInfo) {
            $foundCount++
        }
    }
    
    if ($foundCount -eq $DeviceConfig.Count) {
        $devicesReady = $true
        Write-Log -Message "All devices detected after $elapsedSeconds seconds" -Level SUCCESS
        break
    }
    
    Start-Sleep -Seconds $CheckIntervalSeconds
    $elapsedSeconds += $CheckIntervalSeconds
    
    if ($elapsedSeconds % 10 -eq 0) {
        Write-Log -Message "Still waiting... ($elapsedSeconds seconds elapsed, $foundCount/$($DeviceConfig.Count) devices found)" -Level INFO
    }
}

if (-not $devicesReady) {
    Write-Log -Message "Timeout: Not all devices were detected within $MaxWaitSeconds seconds" -Level WARNING
    Write-Log -Message "Proceeding with available devices..." -Level INFO
}

# Process each device
$successCount = 0
$failCount = 0

foreach ($config in $DeviceConfig) {
    if (Process-DeviceAssignment -Config $config) {
        $successCount++
    }
    else {
        $failCount++
    }
}

# Summary
Write-Log -Message "========== Script Summary ==========" -Level INFO
Write-Log -Message "Successfully configured: $successCount device(s)" -Level INFO
Write-Log -Message "Failed: $failCount device(s)" -Level INFO

if ($failCount -gt 0) {
    Write-Log -Message "Some devices were not configured correctly. Check the log for details." -Level WARNING
}
else {
    Write-Log -Message "All devices configured successfully!" -Level SUCCESS
}

# Continuous monitoring mode (if enabled)
if ($ContinuousMonitoring) {
    Write-Log -Message "Entering continuous monitoring mode..." -Level INFO
    Write-Log -Message "Press Ctrl+C to exit" -Level INFO
    
    while ($true) {
        Start-Sleep -Seconds 30
        
        foreach ($config in $DeviceConfig) {
            $deviceInfo = Get-USBDeviceCOMPort -TargetVendorID $config.VendorID -TargetProductID $config.ProductID
            
            if ($null -ne $deviceInfo) {
                if ($deviceInfo.COMPort -ne $config.TargetCOMPort) {
                    Write-Log -Message "Device drift detected: $($config.Name) on $($deviceInfo.COMPort) instead of $($config.TargetCOMPort)" -Level WARNING
                    Process-DeviceAssignment -Config $config
                }
            }
        }
    }
}

Write-Log -Message "========== USB COM Port Assignment Script Completed ==========" -Level INFO

#endregion
