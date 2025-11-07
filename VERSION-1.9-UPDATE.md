# Version 1.9 Update - Critical COM Settings Persistence Fix

**Release Date:** November 6, 2025  
**Priority:** CRITICAL - Recommended for immediate deployment  
**Compatibility:** 100% backward compatible with Version 1.8

## Critical Issues Fixed

### Issue #1: COM Settings Not Persisting in Device Manager

**Problem:** COM port assignments were working correctly, but Device Manager continued to show default COM settings (9600 baud, etc.) instead of the configured manufacturer specifications.

**Root Cause:** Windows stores COM port settings in **two separate registry locations**:

1. **Device-specific parameters** (HKLM:\SYSTEM\CurrentControlSet\Enum\{DeviceID}\Device Parameters)
   - Individual DWORD values (BaudRate, Parity, StopBits, etc.)
   - Used by the device driver during operation
   
2. **Windows Ports registry** (HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports)
   - Single string value per COM port (e.g., "19200,o,8,1")
   - **What Device Manager displays to users**

Version 1.8 only wrote to location #1, so while devices functioned with correct settings, Device Manager showed incorrect defaults from location #2.

**Solution:** Version 1.9 now writes COM settings to **both** registry locations, ensuring:
- ✅ Device operates with correct settings
- ✅ Device Manager displays correct settings
- ✅ Settings persist across reboots and device reconnections

### Issue #2: Driver Overwriting COM Settings on Restart

**Problem:** When devices required COM port reassignment, COM settings were applied before device restart, causing the USB-to-serial driver to overwrite them with defaults during reinitialization.

**Root Cause:** Incorrect timing sequence:
```
OLD (Wrong) Sequence:
1. Apply COM settings to registry
2. Restart device
3. Driver initializes → Reads .inf defaults → OVERWRITES our settings ❌
```

**Solution:** Version 1.9 implements correct timing logic:

**Scenario A - Port Already Correct (Most Common):**
```
1. Detect device is on correct COM port
2. Apply COM settings directly (no restart needed)
3. Write to both registry locations
4. Done ✅
```

**Scenario B - Port Needs Reassignment:**
```
1. Detect device needs port change
2. Change port assignment in registry
3. Restart device to apply port change
4. Wait 5 seconds for driver to fully initialize
5. Apply COM settings AFTER driver is loaded
6. Write to both registry locations
7. Done ✅
```

## Technical Details

### Dual Registry Write Implementation

```powershell
# Location 1: Device Parameters (device-specific settings)
Set-ItemProperty -Path $deviceParamsPath -Name "BaudRate" -Value 19200 -Type DWord -Force
Set-ItemProperty -Path $deviceParamsPath -Name "Parity" -Value 1 -Type DWord -Force
# ... additional settings

# Location 2: Windows Ports registry (for Device Manager)
$portsPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports"
$portsValue = "19200,o,8,1"  # Format: BaudRate,Parity,DataBits,StopBits
Set-ItemProperty -Path $portsPath -Name "COM5:" -Value $portsValue -Type String -Force
```

### Registry Value Conversions

**Parity Mapping** (Number → Letter):
| Number | Letter | Meaning |
|--------|--------|---------|
| 0 | n | None |
| 1 | o | Odd |
| 2 | e | Even |
| 3 | m | Mark |
| 4 | s | Space |

**Stop Bits Mapping** (Number → Display Value):
| Number | Value | Meaning |
|--------|-------|---------|
| 0 | 1 | 1 stop bit |
| 1 | 1.5 | 1.5 stop bits |
| 2 | 2 | 2 stop bits |

### Device-Specific COM Settings

**Topaz T-LBK462-BSB-RC:**
- Input format: `"19200,8,1,0,0,1,14,14"`
- Device Parameters: BaudRate=19200, DataBits=8, Parity=1 (Odd), StopBits=0 (1 bit)
- Ports value: `"19200,o,8,1"`

**Ingenico LANE3000:**
- Input format: `"115200,8,0,0,0,1,14,14"`
- Device Parameters: BaudRate=115200, DataBits=8, Parity=0 (None), StopBits=0 (1 bit)
- Ports value: `"115200,n,8,1"`

## Upgrade Instructions

### Prerequisites
- Backup current v1.8 script
- Administrative access to VDI gold image
- Test environment available (recommended)

### Deployment Steps

1. **Download Version 1.9**
   ```powershell
   # Download from GitHub
   Invoke-WebRequest -Uri "https://raw.githubusercontent.com/navalynt/usb-com-port-assignment/main/scripts/Set-USBCOMPortAssignment.ps1" `
       -OutFile "C:\Temp\Set-USBCOMPortAssignment-v1.9.ps1"
   ```

2. **Backup Current Script**
   ```powershell
   Copy-Item "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1" `
       -Destination "C:\Imaging\Scripts\Set-USBCOMPortAssignment-v1.8-backup.ps1"
   ```

3. **Deploy New Script**
   ```powershell
   Copy-Item "C:\Temp\Set-USBCOMPortAssignment-v1.9.ps1" `
       -Destination "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1" -Force
   ```

4. **Verify Deployment**
   ```powershell
   # Check script version
   Get-Content "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1" | Select-String "Version"
   ```

### Testing Procedure

**Test Scenario 1: Device Already on Correct Port**
```powershell
# Run script manually
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1

# Expected log output:
# [INFO] Device found on COM5 (Target: COM5)
# [SUCCESS] Device already on correct port: COM5
# [INFO] Applying COM settings without device restart...
# [SUCCESS] Successfully applied COM port settings for COM5
```

**Test Scenario 2: Device Needs Port Reassignment**
```powershell
# Manually assign device to wrong port first, then run script
C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1

# Expected log output:
# [INFO] Device found on COM10 (Target: COM5)
# [INFO] Port reassignment needed: COM10 -> COM5
# [INFO] Restarting device to apply port change...
# [INFO] Waiting 5 seconds for driver initialization...
# [INFO] Driver initialized - applying COM settings...
# [SUCCESS] Successfully applied COM port settings for COM5
```

**Verification: Check Device Manager**
1. Open Device Manager
2. Expand "Ports (COM & LPT)"
3. Right-click Topaz device → Properties → Port Settings

**Expected Results:**
- Bits per second: **19200** (not 9600)
- Data bits: **8**
- Parity: **Odd**
- Stop bits: **1**

**Verification: Check Registry**
```powershell
# Check device-specific settings
$DeviceID = (Get-CimInstance -ClassName Win32_PnPEntity | 
    Where-Object { $_.DeviceID -like "*VID_0403*PID_6001*" }).DeviceID
$RegPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($DeviceID -replace '\\','\\')\Device Parameters"
Get-ItemProperty -Path $RegPath | Select-Object PortName, BaudRate, Parity, StopBits, DataBits

# Check Windows Ports registry
Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports" -Name "COM5:"
```

**Expected Registry Values:**
```
Device Parameters:
  PortName  : COM5
  BaudRate  : 19200
  Parity    : 1
  StopBits  : 0
  DataBits  : 8

Windows Ports:
  COM5:     : 19200,o,8,1
```

## Performance Impact

| Scenario | v1.8 Time | v1.9 Time | Change |
|----------|-----------|-----------|--------|
| Port already correct | ~7 sec | ~7 sec | No change |
| Port needs reassignment | ~12 sec | ~17 sec | +5 sec |

**Note:** The additional 5 seconds in v1.9 for port reassignment is necessary to wait for driver initialization, ensuring COM settings persist correctly.

## Known Issues and Limitations

### None Currently Identified

Version 1.9 resolves all known issues with COM port assignment and settings persistence. If you encounter any problems, please report them with:
- Script version
- Log file excerpt showing error
- Device VID/PID
- Windows version
- Steps to reproduce

## Rollback Procedure

If you need to rollback to Version 1.8:

```powershell
# Restore backup
Copy-Item "C:\Imaging\Scripts\Set-USBCOMPortAssignment-v1.8-backup.ps1" `
    -Destination "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1" -Force

# Verify rollback
Get-Content "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1" | Select-String "Version"
```

## Backward Compatibility

**100% Compatible** - No changes required to:
- ✅ Scheduled tasks
- ✅ DEM FlexEngine wrappers
- ✅ Script parameters
- ✅ Device configuration format
- ✅ COM settings format

## Support and Documentation

- **Full Documentation:** [USB-COM-Port-Assignment-Documentation.md](USB-COM-Port-Assignment-Documentation.md)
- **COM Settings Reference:** [COM-Settings-Quick-Reference.md](COM-Settings-Quick-Reference.md)
- **Deployment Guide:** [USB-COM-Deployment-Guide.md](USB-COM-Deployment-Guide.md)
- **GitHub Repository:** https://github.com/navalynt/usb-com-port-assignment

## Changelog

### Version 1.9 (2025-11-06) - CRITICAL UPDATE
**Added:**
- Dual registry write implementation (Device Parameters + Windows Ports)
- Registry value conversion functions (Parity number → letter, StopBits number → display value)
- Enhanced logging for registry operations

**Fixed:**
- CRITICAL: COM settings now visible in Device Manager
- CRITICAL: Driver initialization timing - settings now persist through device restarts
- Incorrect timing sequence when port reassignment required

**Changed:**
- `Set-COMPortSettings` function now writes to both registry locations
- `Process-DeviceAssignment` function implements correct timing logic
- Port already correct: No restart, immediate COM settings application
- Port needs change: Port change → Restart → Wait 5 sec → Apply COM settings

**Improved:**
- More detailed logging of registry operations
- Better error handling for registry writes
- Clearer log messages indicating timing logic

### Version 1.8 (2025-11-04)
- Fixed registry path resolution
- Enhanced device detection reliability

### Version 1.7 (2025-11-04)
- Added manufacturer-specified COM settings

### Version 1.6 (2025-11-04)
- Added configurable COM port settings

---

**Recommended Action:** Deploy Version 1.9 immediately to all VDI gold images to ensure COM settings persist correctly and display properly in Device Manager.
