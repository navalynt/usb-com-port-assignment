# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.9.0] - 2025-11-06

### 🔴 CRITICAL UPDATE - Recommended for Immediate Deployment

### Added
- **CRITICAL:** Dual registry write implementation for COM settings persistence
  - Now writes to both `Device Parameters` and `Windows Ports` registry locations
  - Ensures Device Manager displays correct COM settings
- Registry value conversion functions
  - Parity: number (0-4) → letter (n/o/e/m/s)
  - Stop Bits: number (0-2) → display value (1/1.5/2)
- Enhanced logging for all registry operations
- Detailed log messages for timing logic decisions

### Fixed
- **CRITICAL:** COM settings now persist correctly after device restarts
  - Fixed driver initialization timing issue
  - Settings no longer overwritten by USB-to-serial driver defaults
- **CRITICAL:** Device Manager now displays configured COM settings
  - Previously showed Windows defaults (9600 baud) instead of configured values
  - Now shows correct manufacturer-specified settings
- Incorrect timing sequence when port reassignment required
  - Old: Apply settings → Restart → Settings lost
  - New: Restart → Wait for driver → Apply settings → Settings persist

### Changed
- `Set-COMPortSettings` function completely rewritten
  - Now writes to both `HKLM:\SYSTEM\CurrentControlSet\Enum\{DeviceID}\Device Parameters` (device-specific)
  - And `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports` (Device Manager display)
  - Added parity and stop bits conversion for Ports registry format
- `Process-DeviceAssignment` function timing logic improved
  - Port already correct: Apply settings immediately without restart (~7 seconds)
  - Port needs change: Change port → Restart → Wait 5 seconds → Apply settings (~17 seconds)
- Enhanced error handling for registry operations
  - Better fallback mechanisms
  - More informative error messages

### Performance
- Port already correct: No change (~7 seconds)
- Port needs reassignment: +5 seconds (~17 seconds total)
  - Additional time ensures driver initialization complete before applying settings

### Documentation
- Added [VERSION-1.9-UPDATE.md](docs/VERSION-1.9-UPDATE.md) with complete upgrade instructions
- Updated README.md with v1.9 features and fixes
- Enhanced technical documentation with dual registry write details

### Backward Compatibility
- ✅ 100% compatible with version 1.8
- ✅ No changes required to scheduled tasks
- ✅ No changes required to DEM FlexEngine wrappers
- ✅ No changes required to script parameters
- ✅ No changes required to device configuration format

---

## [1.8.0] - 2025-11-04

### Fixed
- Registry path resolution in `Set-COMPortSettings` function
  - Now constructs path directly from DeviceID string
  - Added fallback search mechanism if direct construction fails
- Registry path resolution in `Set-COMPortAssignment` function
  - Improved reliability of finding device registry paths

### Changed
- Enhanced registry path detection logic
- Better error messages for registry path failures

---

## [1.7.0] - 2025-11-04

### Changed
- Updated Topaz T-LBK462-BSB-RC COM settings to manufacturer specifications
  - 19200 baud, 8 data bits, Odd parity, 1 stop bit
  - Based on BSB (virtual serial over USB) interface requirements
- Updated Ingenico LANE3000 COM settings to manufacturer specifications
  - 115200 baud, 8 data bits, No parity, 1 stop bit
  - Based on VCOM/USB connection specifications

### Documentation
- Added device-specific COM settings research documentation
- Updated COM Settings Reference Guide with manufacturer specs

---

## [1.6.0] - 2025-11-04

### Added
- Configurable COM port settings support
  - BaudRate, DataBits, Parity, StopBits, FlowControl
  - FIFO buffer settings (RxFIFO, TxFIFO)
- `Set-COMPortSettings` function for applying COM configurations
- COM settings format: "BaudRate,DataBits,Parity,StopBits,FlowControl,UseFIFO,RxBuffer,TxBuffer"
- Support for "default" settings (uses Windows defaults)
- Comprehensive COM settings documentation

### Changed
- Device configuration now includes `COMSettings` parameter
- Enhanced logging to show applied COM settings

### Documentation
- Added [COM-Settings-Quick-Reference.md](docs/COM-Settings-Quick-Reference.md)
- Expanded technical documentation with COM settings examples

---

## [1.5.0] - 2025-11-03

### Added
- Friendly name update in Device Manager
  - Automatically updates to show device name and COM port
  - Example: "Topaz T-LBK462-BSB-RC (COM5)"
- Enhanced user visibility in Device Manager

### Changed
- `Set-COMPortAssignment` function now updates FriendlyName registry value
- Better device identification in Windows

---

## [1.4.0] - 2025-11-03

### Changed
- Renamed VID/PID parameters to VendorID/ProductID throughout solution
  - Resolves PowerShell variable naming conflicts
  - More descriptive parameter names
- Updated all function signatures and documentation

### Fixed
- PowerShell variable scoping issues with short parameter names

---

## [1.3.0] - 2025-11-02

### Added
- Enhanced registry path detection with fallback search mechanism
- More robust device registry path resolution

### Changed
- Improved `Set-COMPortAssignment` function reliability
- Better error handling for registry operations

---

## [1.2.0] - 2025-11-01

### Fixed
- **CRITICAL:** Changed from `Get-CimInstance -Filter` to `Where-Object` filtering
  - WMI filter syntax was failing silently in some environments
  - New approach provides consistent, reliable device detection
  
### Changed
- `Get-USBDeviceCOMPort` function now uses `Where-Object` for filtering
- Improved reliability across different Windows configurations

---

## [1.1.0] - 2025-10-30

### Added
- Continuous monitoring mode with `-ContinuousMonitoring` switch
- Device drift detection (monitors for port changes)
- Configurable check intervals
- Enhanced logging with severity levels

### Changed
- Improved error handling throughout
- Better device detection timing

---

## [1.0.0] - 2025-10-28

### Added
- Initial release
- Automatic COM port assignment for USB devices
- Support for Topaz T-LBK462-BSB-RC signature pad (VID:0403, PID:6001) → COM5
- Support for Ingenico LANE3000 credit card reader (VID:0B00, PID:0084) → COM21
- COM port conflict resolution
- Device restart functionality
- Comprehensive logging system
- Color-coded console output
- UTF-8 encoding support
- PowerShell 5.0 compatibility

### Features
- Device detection based on VendorID/ProductID
- Registry-based COM port assignment
- Automatic conflict resolution (moves conflicting devices to COM100+)
- Device verification after assignment
- Configurable wait times for device initialization
- Non-persistent VDI optimized

### Documentation
- Complete technical documentation
- Deployment guide for Omnissa Horizon View
- Troubleshooting guide
- Examples and best practices

---

## Version Numbering

- **Major version (X.0.0):** Breaking changes or major new features
- **Minor version (0.X.0):** New features, significant improvements, backward compatible
- **Patch version (0.0.X):** Bug fixes, minor improvements, fully backward compatible

## Upgrade Priority Levels

- 🔴 **CRITICAL:** Immediate deployment recommended (fixes major issues)
- 🟡 **RECOMMENDED:** Deploy during next maintenance window (significant improvements)
- 🟢 **OPTIONAL:** Deploy when convenient (minor enhancements)

---

[Unreleased]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.9.0...HEAD
[1.9.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.8.0...v1.9.0
[1.8.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.7.0...v1.8.0
[1.7.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.6.0...v1.7.0
[1.6.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.5.0...v1.6.0
[1.5.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.2.0...v1.3.0
[1.2.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/navalynt/usb-com-port-assignment/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/navalynt/usb-com-port-assignment/releases/tag/v1.0.0
