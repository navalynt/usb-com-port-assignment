# USB COM Port Assignment for Omnissa Horizon VDI

[![PowerShell](https://img.shields.io/badge/PowerShell-5.0%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

> Automated COM port assignment solution for USB devices in non-persistent VDI environments.

## Overview

This PowerShell solution ensures USB devices with COM ports consistently connect to specific COM ports in **Omnissa Horizon View VDI** environments. Designed specifically for non-persistent desktops where devices may connect on different COM ports each session.

### Key Features

- ✅ Automatic COM port assignment based on VID/PID
- ✅ Manufacturer-specified COM settings (baud rate, parity, flow control)
- ✅ Conflict resolution (moves conflicting devices automatically)
- ✅ Device restart to apply changes
- ✅ Friendly name updates in Device Manager
- ✅ Comprehensive logging with color-coded severity
- ✅ Non-persistent VDI optimized
- ✅ PowerShell 5.0 compatible, UTF-8 compliant

### Supported Devices

| Device | Target Port | VID | PID | Settings |
|--------|-------------|-----|-----|----------|
| Topaz T-LBK462-BSB-RC Signature Pad | COM5 | 0403 | 6001 | 19200 baud, Odd parity |
| Ingenico LANE3000 Credit Card Reader | COM21 | 0B00 | 0084 | 115200 baud |

## Quick Start

### Installation

1. **Download the script:**
   ```powershell
   # Download to your gold image
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/YOUR-USERNAME/YOUR-REPO/main/scripts/Set-USBCOMPortAssignment.ps1" `
       -OutFile "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"
   ```

2. **Create scheduled task** (runs as SYSTEM at user logon):
   ```powershell
   $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
       -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"
   
   $trigger = New-ScheduledTaskTrigger -AtLogOn
   
   $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" `
       -LogonType ServiceAccount -RunLevel Highest
   
   $settings = New-ScheduledTaskSettingsSet `
       -AllowStartIfOnBatteries `
       -DontStopIfGoingOnBatteries `
       -StartWhenAvailable
   
   Register-ScheduledTask `
       -TaskName "USB-COM-Port-Assignment" `
       -Action $action `
       -Trigger $trigger `
       -Principal $principal `
       -Settings $settings
   ```

3. **Whitelist devices in Horizon Administrator:**
   - Settings → USB → USB Device Policy
   - Add Topaz: VID `0x0403`, PID `0x6001`
   - Add Ingenico: VID `0x0B00`, PID `0x0084`

4. **Deploy to desktop pools**

### Usage

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

### 📚 Complete Documentation

- **[Technical Documentation](docs/USB-COM-Port-Assignment-Documentation.md)** - Complete technical details
- **[COM Settings Reference](docs/COM-Settings-Quick-Reference.md)** - All COM parameters explained
- **[Deployment Guide](docs/USB-COM-Deployment-Guide.md)** - Step-by-step deployment
- **[Version 1.8 Update](docs/VERSION-1.8-UPDATE.md)** - Critical fix details

### Quick Links

- [Adding New Devices](#adding-new-devices)
- [Troubleshooting](#troubleshooting)
- [COM Settings Format](#com-settings-format)
- [Version History](#version-history)

## Requirements

### System Requirements
- Windows 10/11 (64-bit) or Windows Server 2016/2019/2022+
- PowerShell 5.0 or higher
- Administrator/SYSTEM privileges

### VDI Requirements
- Omnissa Horizon View 2506 (or compatible version)
- USB redirection enabled
- Devices whitelisted in Horizon Administrator

## Configuration

### Device Configuration

Devices are configured in the `$DeviceConfig` array at the top of the script:

```powershell
$DeviceConfig = @(
    @{
        Name = "Topaz T-LBK462-BSB-RC"
        VendorID = "0403"
        ProductID = "6001"
        TargetCOMPort = "COM5"
        COMSettings = "19200,8,1,0,0,1,14,14"  # Manufacturer specifications
    },
    @{
        Name = "Ingenico LANE3000"
        VendorID = "0B00"
        ProductID = "0084"
        TargetCOMPort = "COM21"
        COMSettings = "115200,8,0,0,0,1,14,14"  # Manufacturer specifications
    }
)
```

### COM Settings Format

```
"BaudRate,DataBits,Parity,StopBits,FlowControl,UseFIFO,RxBuffer,TxBuffer"
```

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
   Get-CimInstance Win32_PnPEntity | Where-Object {$_.Name -match 'COM'} | 
       Select-Object Name, DeviceID
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

### Device Not Found

```powershell
# Check USB redirection
Get-Service "VMwareUSBArbitrationService"

# Check Horizon logs
Get-Content "C:\ProgramData\VMware\VDM\logs\vmware-usbarbitrator.log" -Tail 50

# Run diagnostic tool
C:\Imaging\Scripts\Debug-USBCOMDevices.ps1
```

### Wrong COM Port

```powershell
# Check current assignments
Get-CimInstance Win32_PnPEntity | Where-Object {$_.Name -match 'COM'}

# Run script manually
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1

# Check log
Get-Content "C:\Temp\USBCOMPortAssignment_$(Get-Date -Format 'yyyyMMdd').log"
```

### COM Settings Not Applied

```powershell
# Check registry (for Topaz)
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB\VID_0403&PID_6001"
Get-ChildItem $regPath | ForEach-Object {
    Get-ItemProperty "$($_.PSPath)\Device Parameters" -Name BaudRate, RxFIFO, TxFIFO
}

# Should show: BaudRate: 19200, RxFIFO: 14, TxFIFO: 14
```

See [Technical Documentation](docs/USB-COM-Port-Assignment-Documentation.md#troubleshooting) for more.

## Scripts

### Main Scripts

- **[Set-USBCOMPortAssignment.ps1](scripts/Set-USBCOMPortAssignment.ps1)** (v1.8) - Main script for COM port management
- **[Debug-USBCOMDevices.ps1](scripts/Debug-USBCOMDevices.ps1)** (v1.0) - Diagnostic and troubleshooting tool

### Script Features

**Set-USBCOMPortAssignment.ps1:**
- Waits for USB devices to initialize
- Assigns devices to target COM ports
- Applies manufacturer-specified COM settings
- Resolves COM port conflicts
- Restarts devices to apply changes
- Updates friendly names in Device Manager
- Comprehensive logging

**Debug-USBCOMDevices.ps1:**
- Scans all USB and COM devices
- Checks registry assignments
- Tests script logic
- Identifies problems
- Generates detailed diagnostic report

## Logs

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
| **1.8** | **Nov 4, 2025** | **Fixed registry path resolution (CRITICAL)** |

See [VERSION-1.8-UPDATE.md](docs/VERSION-1.8-UPDATE.md) for details on the latest fix.

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Reporting Issues

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

- Built for **Omnissa Horizon View 2506** environments
- Designed for non-persistent VDI with FSLogix, App Volumes, and DEM
- COM settings based on manufacturer specifications:
  - [Topaz Systems](https://www.topazsystems.com) - T-LBK462-BSB-RC specifications
  - [Ingenico](https://www.ingenico.com) - LANE3000 VCOM/USB specifications

## Support

- **Documentation:** See [docs](docs/) folder
- **Issues:** [GitHub Issues](https://github.com/YOUR-USERNAME/YOUR-REPO/issues)
- **Discussions:** [GitHub Discussions](https://github.com/YOUR-USERNAME/YOUR-REPO/discussions)

## Related Projects

- [Omnissa Horizon View Documentation](https://docs.omnissa.com/)
- [PowerShell Gallery](https://www.powershellgallery.com/)

---

**Made with ❤️ for VDI Administrators**

**Current Version:** 1.8 | **PowerShell:** 5.0+ | **Platform:** Windows 10/11, Server 2016+
