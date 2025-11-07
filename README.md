# USB COM Port Assignment for Omnissa Horizon VDI

Automated COM port assignment solution for USB devices in non-persistent VDI environments.

This PowerShell solution ensures USB devices with COM ports consistently connect to specific COM ports in Omnissa Horizon View VDI environments. Designed specifically for non-persistent desktops where devices may connect on different COM ports each session.

## Features

- ✅ Automatic COM port assignment based on VID/PID
- ✅ Manufacturer-specified COM settings (baud rate, parity, flow control)
- ✅ **NEW v1.9:** Dual registry writes for Device Manager visibility
- ✅ **NEW v1.9:** Intelligent timing logic to prevent driver overwrites
- ✅ Conflict resolution (moves conflicting devices automatically)
- ✅ Device restart to apply changes
- ✅ Friendly name updates in Device Manager
- ✅ Comprehensive logging with color-coded severity
- ✅ Non-persistent VDI optimized
- ✅ PowerShell 5.0 compatible, UTF-8 compliant

## Supported Devices

| Device | Target Port | VID | PID | Settings |
|--------|------------|-----|-----|----------|
| Topaz T-LBK462-BSB-RC Signature Pad | COM5 | 0403 | 6001 | 19200 baud, Odd parity |
| Ingenico LANE3000 Credit Card Reader | COM21 | 0B00 | 0084 | 115200 baud |

## Quick Start

### 1. Download the script

```powershell
# Download to your gold image
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/navalynt/usb-com-port-assignment/main/scripts/Set-USBCOMPortAssignment.ps1" `
    -OutFile "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"
```

### 2. Create scheduled task (runs as SYSTEM at user logon)

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Minimized -File C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"

$trigger = New-ScheduledTaskTrigger -AtLogOn

$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" `
    -LogonType ServiceAccount -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -Priority 3

Register-ScheduledTask `
    -TaskName "USB-COM-Port-Assignment" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings
```

**Or use the automated installer:**

```powershell
# Download installer
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/navalynt/usb-com-port-assignment/main/scripts/Install-COMPortAssignmentTask.ps1" `
    -OutFile "C:\Temp\Install-COMPortAssignmentTask.ps1"

# Run as Administrator
.\Install-COMPortAssignmentTask.ps1
```

### 3. Whitelist devices in Horizon DEM Computer Policy Management Console

- Computer Environment → Horizon Smart Policies → USB Device Policy (create if doesn't exist)
- **Exclude all devices** should be ENABLED for security reasons
- Add the below devices to "Include VID/PID device" line, semi-colon separated
  - Add Topaz: `VID-0403_PID-6001`
  - Add Ingenico: `VID-0B00_PID-0084`
  - Ex: `o:VID-0403_PID-6001;VID-0B00_PID-0084`
  - **"o"** overrides any settings on the client and is therefore preferred

[Omnissa Documentation - USB Device Policies](https://docs.omnissa.com/bundle/Horizon-Remote-Desktop-FeaturesVmulti/page/ConfiguringFilterPolicySettingsforUSBDevices.html)

### 4. Deploy to desktop pools

The script runs automatically at user logon. To run manually:

```powershell
# Standard execution
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1

# With extended wait time
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1 -MaxWaitSeconds 120

# With continuous monitoring
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1 -ContinuousMonitoring
```

## Documentation

- [Technical Documentation](docs/USB-COM-Port-Assignment-Documentation.md) - Complete technical details
- [COM Settings Reference](docs/COM-Settings-Quick-Reference.md) - All COM parameters explained
- [Deployment Guide](docs/USB-COM-Deployment-Guide.md) - Step-by-step deployment
- [**Version 1.9 Update**](docs/VERSION-1.9-UPDATE.md) - **CRITICAL fix details - READ THIS FIRST**

## What's New in Version 1.9

### CRITICAL FIXES

**Issue #1: COM Settings Not Visible in Device Manager** ✅ FIXED
- Device Manager now correctly displays configured COM settings
- Dual registry write implementation ensures persistence

**Issue #2: Driver Overwriting Settings on Restart** ✅ FIXED
- Intelligent timing logic prevents driver from overwriting settings
- Correct sequence: Port change → Restart → Wait for driver → Apply settings

See [VERSION-1.9-UPDATE.md](docs/VERSION-1.9-UPDATE.md) for complete details and upgrade instructions.

## Requirements

- Windows 10/11 (64-bit) or Windows Server 2016/2019/2022+
- PowerShell 5.0 or higher
- Administrator/SYSTEM privileges
- Omnissa Horizon View 2506 (or compatible version)
- USB redirection enabled
- Devices whitelisted in Horizon Administrator

## Configuration

Devices are configured in the `$DeviceConfig` array at the top of the script:

```powershell
$DeviceConfig = @(
    @{
        Name = "Topaz T-LBK462-BSB-RC"
        VendorID = "0403"
        ProductID = "6001"
        TargetCOMPort = "COM5"
        COMSettings = "19200,8,1,0,0,1,14,14" # Manufacturer specifications
    },
    @{
        Name = "Ingenico LANE3000"
        VendorID = "0B00"
        ProductID = "0084"
        TargetCOMPort = "COM21"
        COMSettings = "115200,8,0,0,0,1,14,14" # Manufacturer specifications
    }
)
```

### COM Settings Format

`"BaudRate,DataBits,Parity,StopBits,FlowControl,UseFIFO,RxBuffer,TxBuffer"`

| Position | Parameter | Values | Description |
|----------|-----------|--------|-------------|
| 1 | Baud Rate | 9600, 19200, 115200, etc. | Communication speed |
| 2 | Data Bits | 5, 6, 7, 8 | Usually 8 |
| 3 | Parity | 0=None, 1=Odd, 2=Even | Error checking |
| 4 | Stop Bits | 0=1bit, 1=1.5bits, 2=2bits | Usually 0 (1 bit) |
| 5 | Flow Control | 0=None, 1=Xon/Xoff, 2=Hardware | Usually 0 |
| 6 | Use FIFO | 0=Disabled, 1=Enabled | Always use 1 |
| 7 | RX Buffer | 1-14 | Receive buffer size (14=max) |
| 8 | TX Buffer | 1-14 | Transmit buffer size (14=max) |

Use `"default"` to skip custom settings and use Windows defaults.

## Adding New Devices

1. **Find the device VID/PID:**
   ```powershell
   Get-CimInstance Win32_PnPEntity | Where-Object {$_.Name -match 'COM'} | Select-Object Name, DeviceID
   ```

2. **Add to the configuration:**
   ```powershell
   @{
       Name = "Your Device Name"
       VendorID = "XXXX"  # From step 1
       ProductID = "YYYY"  # From step 1
       TargetCOMPort = "COMZ"  # Your choice
       COMSettings = "default"  # Or custom settings
   }
   ```

3. **Whitelist in Horizon Administrator**

4. **Test and deploy**

See [Technical Documentation](docs/USB-COM-Port-Assignment-Documentation.md#adding-new-devices) for details.

## Troubleshooting

### Check USB Redirection

```powershell
# Check USB redirection
Get-Service "VMUSBArbService"

# Check Horizon logs
Get-Content "C:\ProgramData\Omnissa\Horizon\logs\vmware-usbarbitrator.log" -Tail 50
```

### Run Diagnostic Tool

```powershell
# Run diagnostic tool
C:\Imaging\Scripts\Debug-USBCOMDevices.ps1

# Check current assignments
Get-CimInstance Win32_PnPEntity | Where-Object {$_.Name -match 'COM'}
```

### Manual Test Run

```powershell
# Run script manually
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1

# Check log
Get-Content "C:\Temp\USBCOMPortAssignment_$(Get-Date -Format 'yyyyMMdd').log"
```

### Verify COM Settings

```powershell
# Check registry (for Topaz)
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\VID_0403&PID_6001"
Get-ChildItem $regPath | ForEach-Object {
    Get-ItemProperty "$($_.PSPath)\Device Parameters" -Name BaudRate, RxFIFO, TxFIFO
}

# Should show: BaudRate: 19200, RxFIFO: 14, TxFIFO: 14

# Check Windows Ports registry (for Device Manager visibility)
Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" -Name "COM5:"

# Should show: COM5: = 19200,o,8,1
```

See [Technical Documentation](docs/USB-COM-Port-Assignment-Documentation.md#troubleshooting) for more.

## Scripts

- [Set-USBCOMPortAssignment.ps1](scripts/Set-USBCOMPortAssignment.ps1) (v1.9) - Main script for COM port management
- [Install-COMPortAssignmentTask.ps1](scripts/Install-COMPortAssignmentTask.ps1) (v1.1) - Automated scheduled task installer
- [Start-COMPortAssignment-DEM.ps1](scripts/Start-COMPortAssignment-DEM.ps1) (v1.0) - DEM FlexEngine wrapper
- [Debug-USBCOMDevices.ps1](scripts/Debug-USBCOMDevices.ps1) (v1.0) - Diagnostic and troubleshooting tool

### Main Script Functions

**Set-USBCOMPortAssignment.ps1:**
- Waits for USB devices to initialize
- Assigns devices to target COM ports
- Applies manufacturer-specified COM settings with dual registry writes
- Resolves COM port conflicts
- Restarts devices to apply changes (with intelligent timing)
- Updates friendly names in Device Manager
- Comprehensive logging

**Debug-USBCOMDevices.ps1:**
- Scans all USB and COM devices
- Checks registry assignments
- Tests script logic
- Identifies problems
- Generates detailed diagnostic report

## Logging

Logs are written to `C:\Temp\USBCOMPortAssignment_YYYYMMDD.log`

```powershell
# View today's log
Get-Content "C:\Temp\USBCOMPortAssignment_$(Get-Date -Format 'yyyyMMdd').log"

# Check for errors
Get-Content "C:\Temp\USBCOMPortAssignment_*.log" | Select-String "\[ERROR\]"
```

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | Oct 2025 | Initial release |
| 1.6 | Nov 2025 | Added configurable COM settings |
| 1.7 | Nov 4, 2025 | Updated to manufacturer specifications |
| 1.8 | Nov 4, 2025 | Fixed registry path resolution (CRITICAL) |
| **1.9** | **Nov 6, 2025** | **Dual registry writes + timing fixes (CRITICAL)** |

See [VERSION-1.9-UPDATE.md](docs/VERSION-1.9-UPDATE.md) for details on the latest critical fixes.

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

When reporting issues, please include:
- Script version
- Windows version
- PowerShell version
- Device VID/PID
- Log file excerpt showing error
- Steps to reproduce

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built for Omnissa Horizon View 2506 environments
- Designed for non-persistent VDI with FSLogix, App Volumes, and DEM
- COM settings based on manufacturer specifications:
  - [Topaz Systems](https://www.topazsystems.com) - T-LBK462-BSB-RC specifications
  - [Ingenico](https://www.ingenico.com) - LANE3000 VCOM/USB specifications

## Support

- **Documentation:** See [docs](docs) folder
- **Issues:** [GitHub Issues](https://github.com/navalynt/usb-com-port-assignment/issues)
- **Discussions:** [GitHub Discussions](https://github.com/navalynt/usb-com-port-assignment/discussions)

---

**Made with ❤️ for VDI Administrators**

**Current Version:** 1.9 | **PowerShell:** 5.0+ | **Platform:** Windows 10/11, Server 2016+
