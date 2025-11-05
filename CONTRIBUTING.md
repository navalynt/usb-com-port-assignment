# Contributing to USB COM Port Assignment

First off, thank you for considering contributing to this project! It's people like you that make this tool better for the VDI community.

## Code of Conduct

This project and everyone participating in it is governed by a professional code of conduct. By participating, you are expected to uphold this code.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the existing issues to avoid duplicates.

When you create a bug report, include as many details as possible:

- **Use a clear and descriptive title**
- **Describe the exact steps to reproduce the problem**
- **Provide specific examples** (device VID/PID, Windows version, etc.)
- **Describe the behavior you observed** and what you expected
- **Include screenshots or log files** if applicable
- **Environment details:**
  - Script version
  - Windows version
  - PowerShell version
  - VDI platform and version
  - Device information (VID/PID, model)

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion:

- **Use a clear and descriptive title**
- **Provide a detailed description** of the suggested enhancement
- **Provide specific examples** of how it would be used
- **Explain why this enhancement would be useful**

### Adding Support for New Devices

One of the most valuable contributions is adding support for new USB COM devices:

1. **Test the device** with the existing script
2. **Document the device:**
   - Manufacturer and model
   - VID and PID
   - Optimal COM settings (from manufacturer documentation)
3. **Submit configuration:**
   ```powershell
   @{
       Name = "Device Name"
       VendorID = "XXXX"
       ProductID = "YYYY"
       TargetCOMPort = "COMZ"
       COMSettings = "baud,data,parity,stop,flow,fifo,rx,tx"
   }
   ```
4. **Include testing results** in your pull request

### Pull Requests

1. **Fork the repository** and create your branch from `main`
2. **Follow the PowerShell style guide** (see below)
3. **Update documentation** if you change functionality
4. **Test your changes** thoroughly
5. **Update CHANGELOG.md** with your changes
6. **Ensure UTF-8 encoding** for all files
7. **Submit the pull request**

## Development Guidelines

### PowerShell Style Guide

This project follows these PowerShell standards:

1. **Encoding:**
   - Always use UTF-8 encoding
   - No BOM (Byte Order Mark)
   - No special characters or emoji in code

2. **Variable Naming:**
   - Use PascalCase for functions: `Get-DeviceInfo`
   - Use camelCase for variables: `$devicePath`
   - Avoid reserved variable names (VID → VendorID, PID → ProductID)
   - Use descriptive names: `$targetCOMPort` not `$tp`

3. **Functions:**
   - Use `[CmdletBinding()]` attribute
   - Define parameters with proper types
   - Include parameter validation where appropriate
   - Add comment-based help

4. **Error Handling:**
   - Use try/catch blocks
   - Log all errors with context
   - Provide meaningful error messages
   - Never silently fail

5. **Logging:**
   - Use the `Write-Log` function
   - Include appropriate severity levels (INFO, WARNING, ERROR, SUCCESS)
   - Log significant actions and their results

6. **Comments:**
   - Use comments for complex logic
   - Explain "why" not "what"
   - Keep comments up to date

### Example Function Template

```powershell
function Verb-Noun {
    <#
    .SYNOPSIS
        Brief description
    
    .DESCRIPTION
        Detailed description
    
    .PARAMETER ParameterName
        Description of parameter
    
    .EXAMPLE
        Verb-Noun -ParameterName "value"
        Description of example
    
    .NOTES
        Additional notes
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ParameterName
    )
    
    try {
        Write-Log -Message "Starting Verb-Noun operation" -Level INFO
        
        # Your code here
        
        Write-Log -Message "Successfully completed Verb-Noun" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log -Message "Error in Verb-Noun: $_" -Level ERROR
        return $false
    }
}
```

### Testing

Before submitting a pull request:

1. **Test on PowerShell 5.0** (minimum version)
2. **Test on Windows 10 and Windows 11**
3. **Test with actual devices** if possible
4. **Test in VDI environment** if available
5. **Check logs** for any errors or warnings
6. **Verify UTF-8 encoding:**
   ```powershell
   $content = Get-Content "YourScript.ps1" -Raw
   [System.Text.Encoding]::UTF8.GetBytes($content)
   ```

### Documentation

When adding features or making changes:

1. **Update README.md** if user-facing changes
2. **Update technical documentation** in docs/
3. **Update CHANGELOG.md** with your changes
4. **Add examples** for new functionality
5. **Update COM settings reference** if adding device types

## Project Structure

```
.
├── README.md                           # Main documentation
├── LICENSE                             # MIT License
├── CHANGELOG.md                        # Version history
├── CONTRIBUTING.md                     # This file
├── .gitignore                          # Git ignore rules
├── scripts/
│   ├── Set-USBCOMPortAssignment.ps1   # Main script
│   └── Debug-USBCOMDevices.ps1        # Diagnostic tool
└── docs/
    ├── USB-COM-Port-Assignment-Documentation.md
    ├── COM-Settings-Quick-Reference.md
    ├── USB-COM-Deployment-Guide.md
    └── VERSION-1.8-UPDATE.md
```

## Commit Message Guidelines

Use clear and meaningful commit messages:

- **feat:** New feature
- **fix:** Bug fix
- **docs:** Documentation changes
- **style:** Formatting changes (no code change)
- **refactor:** Code refactoring
- **test:** Adding tests
- **chore:** Maintenance tasks

Examples:
```
feat: Add support for Zebra ZD420 printer
fix: Correct registry path resolution in Set-COMPortSettings
docs: Update deployment guide with new prerequisites
```

## Questions?

Feel free to:
- Open an issue with the `question` label
- Start a discussion in GitHub Discussions
- Contact the maintainers

## Recognition

Contributors will be recognized in:
- GitHub Contributors page
- Project README (for significant contributions)
- Release notes

Thank you for contributing! 🎉
