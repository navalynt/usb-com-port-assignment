#Requires -Version 5.0
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs scheduled task for USB COM Port Assignment script.

.DESCRIPTION
    Creates a scheduled task that runs the USB COM Port Assignment script
    at user logon. The task runs as SYSTEM account with highest privileges.

.NOTES
    File Name      : Install-COMPortAssignmentTask.ps1
    Author         : VDI Admin
    Prerequisite   : PowerShell 5.0, Run as Administrator
    Version        : 1.1
    Location       : C:\Imaging\Scripts\
    
.EXAMPLE
    .\Install-COMPortAssignmentTask.ps1
    
.EXAMPLE
    # Specify custom script location
    .\Install-COMPortAssignmentTask.ps1 -ScriptPath "D:\Scripts\Set-USBCOMPortAssignment.ps1"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ScriptPath = "C:\Imaging\Scripts\Set-USBCOMPortAssignment.ps1",
    
    [Parameter(Mandatory=$false)]
    [string]$TaskName = "USB-COM-Port-Assignment"
)

# Set output encoding to UTF-8
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # Silently continue if no console is available
}
$OutputEncoding = [System.Text.Encoding]::UTF8

function Write-ColorOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White
    )
    
    try {
        Write-Host $Message -ForegroundColor $ForegroundColor
    } catch {
        Write-Output $Message
    }
}

# Check if running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-ColorOutput "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    exit 1
}

Write-ColorOutput "`n========================================" -ForegroundColor Cyan
Write-ColorOutput "USB COM Port Assignment Task Installer" -ForegroundColor Cyan
Write-ColorOutput "========================================`n" -ForegroundColor Cyan

# Verify script exists
if (-not (Test-Path -Path $ScriptPath)) {
    Write-ColorOutput "ERROR: Script not found at: $ScriptPath" -ForegroundColor Red
    Write-ColorOutput "Please ensure the script is in the correct location." -ForegroundColor Yellow
    exit 1
}

Write-ColorOutput "Script location: $ScriptPath" -ForegroundColor Green
Write-ColorOutput "Task name: $TaskName`n" -ForegroundColor Green

# Check if task already exists
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($null -ne $existingTask) {
    Write-ColorOutput "WARNING: Task '$TaskName' already exists." -ForegroundColor Yellow
    $response = Read-Host "Do you want to remove and recreate it? (Y/N)"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-ColorOutput "Removing existing task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-ColorOutput "Existing task removed." -ForegroundColor Green
    }
    else {
        Write-ColorOutput "Installation cancelled." -ForegroundColor Yellow
        exit 0
    }
}

try {
    Write-ColorOutput "Creating scheduled task..." -ForegroundColor Cyan
    
    # Task Action - Run PowerShell with the script
    # -WindowStyle Minimized: Shows window minimized instead of hidden for visibility
    $action = New-ScheduledTaskAction `
        -Execute "PowerShell.exe" `
        -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Minimized -File `"$ScriptPath`""
    
    Write-ColorOutput "  Action configured: PowerShell with minimized window" -ForegroundColor Gray
    
    # Task Trigger - At user logon
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    
    Write-ColorOutput "  Trigger configured: At user logon" -ForegroundColor Gray
    
    # Task Principal - Run as SYSTEM with highest privileges
    $principal = New-ScheduledTaskPrincipal `
        -UserId "SYSTEM" `
        -LogonType ServiceAccount `
        -RunLevel Highest
    
    Write-ColorOutput "  Principal configured: SYSTEM account with highest privileges" -ForegroundColor Gray
    
    # Task Settings
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -Priority 3 `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    
    Write-ColorOutput "  Settings configured: Priority 3, battery-friendly, 10-minute timeout" -ForegroundColor Gray
    
    # Register the task
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Principal $principal `
        -Settings $settings `
        -Description "Assigns static COM ports to USB devices (Topaz signature pad to COM5, Ingenico credit card reader to COM21) in VDI environment." | Out-Null
    
    Write-ColorOutput "`nSUCCESS: Scheduled task '$TaskName' created successfully!" -ForegroundColor Green
    
    # Display task details
    Write-ColorOutput "`nTask Details:" -ForegroundColor Cyan
    Write-ColorOutput "  Name:          $TaskName" -ForegroundColor White
    Write-ColorOutput "  Status:        Ready" -ForegroundColor White
    Write-ColorOutput "  Run As:        SYSTEM" -ForegroundColor White
    Write-ColorOutput "  Trigger:       At user logon" -ForegroundColor White
    Write-ColorOutput "  Priority:      3 (High)" -ForegroundColor White
    Write-ColorOutput "  Window Style:  Minimized" -ForegroundColor White
    Write-ColorOutput "  Script:        $ScriptPath" -ForegroundColor White
    
    # Test run option
    Write-ColorOutput "`nWould you like to test run the task now? (Y/N)" -ForegroundColor Yellow
    $testRun = Read-Host
    
    if ($testRun -eq 'Y' -or $testRun -eq 'y') {
        Write-ColorOutput "`nStarting task..." -ForegroundColor Cyan
        Start-ScheduledTask -TaskName $TaskName
        
        Write-ColorOutput "Task started. Check C:\Temp\USBCOMPortAssignment_*.log for results." -ForegroundColor Green
        Write-ColorOutput "Note: The PowerShell window will appear minimized on the taskbar." -ForegroundColor Yellow
    }
    
    Write-ColorOutput "`nInstallation complete!" -ForegroundColor Green
    Write-ColorOutput "The task will run automatically at user logon.`n" -ForegroundColor White
}
catch {
    Write-ColorOutput "`nERROR: Failed to create scheduled task!" -ForegroundColor Red
    Write-ColorOutput "Error details: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
