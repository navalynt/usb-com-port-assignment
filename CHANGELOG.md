# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.8.0] - 2025-11-04

### Fixed
- **CRITICAL:** Fixed registry path resolution in `Set-COMPortSettings` and `Set-COMPortAssignment` functions
- COM settings now correctly apply (baud rate, parity, flow control, buffers)
- Registry path is now constructed directly from DeviceID instead of searching for non-existent property

### Changed
- Improved logging in registry path resolution
- Added fallback search method for registry paths

### Impact
- Version 1.7 and earlier: COM port assignments worked, but COM settings were not applied
- Version 1.8: Both COM port assignments AND COM settings now work correctly
- Devices now use manufacturer-specified settings instead of Windows defaults

## [1.7.0] - 2025-11-04

### Changed
- Updated device name from "Topaz Signature Pad" to "Topaz T-LBK462-BSB-RC" (specific model)
- Updated Topaz COM settings to manufacturer specifications: `19200,8,1,0,0,1,14,14`
  - Baud rate: 19200 (BSB specification)
  - Parity: Odd (manufacturer requirement)
- Confirmed Ingenico COM settings: `115200,8,0,0,0,1,14,14`
  - Baud rate: 115200 (manufacturer maximum for VCOM/USB)

### Documentation
- Enhanced documentation with manufacturer research
- Added device-specific notes and specifications

## [1.6.0] - 2025-11-03

### Added
- Configurable COM port settings (baud rate, data bits, parity, stop bits, flow control, FIFO, buffers)
- New `Set-COMPortSettings` function to apply custom COM settings
- Support for "default" settings option
- Comprehensive COM settings documentation

### Documentation
- Added COM-Settings-Quick-Reference.md
- Detailed explanation of all 8 COM parameters
- Real-world examples for multiple device types
- Troubleshooting by symptom guide

## [1.5.0] - 2025-11-02

### Added
- Friendly name update in Device Manager
- Devices now display as "Device Name (COMX)" format

### Changed
- Improved user experience with clear device identification
- Enhanced Device Manager integration

## [1.4.0] - 2025-11-01

### Changed
- Renamed VID/PID variables to VendorID/ProductID throughout script
- Resolved PowerShell reserved variable conflicts
- Improved variable naming consistency

### Fixed
- PowerShell variable name conflicts that could cause unexpected behavior

## [1.3.0] - 2025-10-31

### Changed
- Enhanced registry path detection with fallback search mechanism
- More reliable device parameter access
- Improved error handling for registry operations

### Fixed
- Registry path detection failures in some scenarios

## [1.2.0] - 2025-10-30

### Changed
- Changed from WMI `-Filter` parameter to `Where-Object` cmdlet
- Better PowerShell version compatibility
- More reliable device queries

### Fixed
- Compatibility issues with certain PowerShell configurations

## [1.0.0] - 2025-10-29

### Added
- Initial release
- Automatic COM port assignment based on VID/PID
- Support for Topaz T-LBK462-BSB-RC signature pad (COM5)
- Support for Ingenico LANE3000 credit card reader (COM21)
- Device restart functionality
- Comprehensive logging
- Wait logic for device initialization
- Conflict resolution (moves conflicting devices)
- Scheduled task deployment support
- Diagnostic tool (Debug-USBCOMDevices.ps1)

### Features
- PowerShell 5.0 compatible
- UTF-8 compliant
- Non-persistent VDI optimized
- Omnissa Horizon View 2506 compatible
- FSLogix, App Volumes, and DEM integration

### Documentation
- Complete technical documentation
- Deployment guide
- Quick start guide
- COM settings reference

---

## Version Numbering

This project uses [Semantic Versioning](https://semver.org/):
- MAJOR version for incompatible API changes
- MINOR version for added functionality in a backwards compatible manner
- PATCH version for backwards compatible bug fixes

## Upgrade Notes

### Upgrading to 1.8.0 from 1.7.0 or earlier
**IMPORTANT:** This is a critical bug fix. Previous versions assigned COM ports correctly but did not apply custom COM settings.

**Action Required:**
1. Replace script on gold image
2. Test on single desktop to verify COM settings now apply
3. Check log for "Successfully applied COM port settings" messages
4. Verify registry shows correct BaudRate, RxFIFO, TxFIFO values
5. Deploy to all pools

**Expected Changes:**
- Devices will now use manufacturer-specified settings
- Topaz pad runs at 19200 baud with Odd parity
- Ingenico terminal runs at 115200 baud
- Better device performance and reliability

### Upgrading to 1.7.0 from 1.6.0 or earlier
**Action Required:**
1. Update device configurations with manufacturer specifications
2. Test devices after deployment
3. Verify applications work with new settings

### Upgrading to 1.6.0 from 1.5.0 or earlier
**Action Required:**
1. Review COM settings documentation
2. Determine if custom settings are needed for your devices
3. Update `$DeviceConfig` array if needed
4. Test thoroughly before production deployment

## Links

- [Latest Release](https://github.com/YOUR-USERNAME/YOUR-REPO/releases/latest)
- [All Releases](https://github.com/YOUR-USERNAME/YOUR-REPO/releases)
- [Issue Tracker](https://github.com/YOUR-USERNAME/YOUR-REPO/issues)
