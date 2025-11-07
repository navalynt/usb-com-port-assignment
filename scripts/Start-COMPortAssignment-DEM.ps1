#Requires -Version 5.0
<#
.SYNOPSIS
    DEM FlexEngine wrapper for USB COM Port Assignment script.

.DESCRIPTION
    This wrapper script is designed to be called from Omnissa DEM FlexEngine
    computer policy. It launches the main USB COM Port Assignment script
    with appropriate logging for DEM integration.

.NOTES
    File Name      : Start-COMPortAssignment-DEM.ps1
    Author         : VDI Admin
    Prerequisite   : PowerShell 5.0, Omnissa DEM with computer policy enabled
    Version        : 1.0
    Location       : C:\Imaging\Scripts\
    
.EXAMPLE
    # Called from DEM FlexEngine Computer Policy
    PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File "C:\Imaging\Scripts\Start-COMPortAssignment-DEM.ps1"

.NOTES
    DEM FlexEngine Configuration:
    1. Open DEM Management Console
    2. Navigate to Computer Policy
    3. Create new FlexEngine Computer Assignment
    4. Name: USB COM Port Assignment
    5. Executable: PowerShell.exe
    6. Arguments: -ExecutionPolicy Bypass -NoProfile -WindowStyle Minimized -File "C:\Imaging\Scripts\Start-COMPortAssignment-DEM.ps1"
    7. Run as: SYSTEM
    8. Trigger: At computer startup or user logon
#>

[CmdletBinding()]
param()

# Set output encoding to UTF-8
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # Silently continue if no console is available
}
$OutputEncoding = [System.Text.Encoding]::UTF8

# Configuration
$MainScriptPath = "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1"
$LogPath = "C:\Temp"
$WrapperLogFile = Join-Path -Path $LogPath -ChildPath "DEM-USBCOMPortAssignment_$(Get-Date -Format 'yyyyMMdd').log"

function Write-WrapperLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [DEM-WRAPPER] [$Level] $Message"
    
    # Ensure log directory exists
    if (-not (Test-Path -Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    # Write to log file
    Add-Content -Path $WrapperLogFile -Value $logMessage -Encoding UTF8
}

Write-WrapperLog -Message "========== DEM FlexEngine Wrapper Started ==========" -Level INFO
Write-WrapperLog -Message "Computer: $env:COMPUTERNAME" -Level INFO
Write-WrapperLog -Message "User Context: $env:USERNAME" -Level INFO
Write-WrapperLog -Message "Main Script: $MainScriptPath" -Level INFO

# Verify main script exists
if (-not (Test-Path -Path $MainScriptPath)) {
    Write-WrapperLog -Message "ERROR: Main script not found at: $MainScriptPath" -Level ERROR
    Write-WrapperLog -Message "DEM FlexEngine wrapper cannot proceed without main script." -Level ERROR
    Write-WrapperLog -Message "========== DEM FlexEngine Wrapper Failed ==========" -Level ERROR
    exit 1
}

Write-WrapperLog -Message "Main script found. Launching USB COM Port Assignment..." -Level INFO

try {
    # Execute the main script
    $startTime = Get-Date
    
    & PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File $MainScriptPath
    
    $endTime = Get-Date
    $duration = ($endTime - $startTime).TotalSeconds
    
    Write-WrapperLog -Message "Main script completed in $([math]::Round($duration, 2)) seconds" -Level SUCCESS
    Write-WrapperLog -Message "Exit code: $LASTEXITCODE" -Level INFO
    
    if ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE) {
        Write-WrapperLog -Message "USB COM Port Assignment completed successfully via DEM FlexEngine" -Level SUCCESS
    }
    else {
        Write-WrapperLog -Message "USB COM Port Assignment completed with exit code: $LASTEXITCODE" -Level WARNING
    }
}
catch {
    Write-WrapperLog -Message "ERROR: Exception occurred while running main script" -Level ERROR
    Write-WrapperLog -Message "Exception: $($_.Exception.Message)" -Level ERROR
    Write-WrapperLog -Message "========== DEM FlexEngine Wrapper Failed ==========" -Level ERROR
    exit 1
}

Write-WrapperLog -Message "========== DEM FlexEngine Wrapper Completed ==========" -Level INFO
exit 0
