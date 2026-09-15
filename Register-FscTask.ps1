#Requires -Version 5.1
<#
.SYNOPSIS
    Registers a Windows scheduled task that refreshes the fuel surcharge board daily.

.DESCRIPTION
    Creates a task under your own user account. No administrator rights and no
    software installation are required - Task Scheduler is built into Windows.

    The task runs Get-FuelSurcharges.ps1 once a day. If the machine is off or
    asleep at the scheduled time, it runs at the next opportunity instead of
    skipping the day.

.PARAMETER At
    Time of day to run, 24h. Default 07:15 - after the carriers post Monday changes
    but before the workday starts.

.PARAMETER TaskName
    Name shown in Task Scheduler. Default "ACE Fuel Surcharge Board".

.PARAMETER Unregister
    Remove the task instead of creating it.

.PARAMETER RunNow
    Trigger the task once immediately after registering, to prove it works.

.EXAMPLE
    .\Register-FscTask.ps1 -RunNow
    .\Register-FscTask.ps1 -At 06:30
    .\Register-FscTask.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [string] $At       = '07:15',
    [string] $TaskName = 'ACE Fuel Surcharge Board',
    [switch] $Unregister,
    [switch] $RunNow
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
    throw 'The ScheduledTasks module is not available on this machine. Use Task Scheduler''s UI instead - see README.md.'
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($Unregister) {
    if ($existing) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
    }
    else {
        Write-Host "No scheduled task named '$TaskName' to remove." -ForegroundColor DarkGray
    }
    return
}

$script = Join-Path $PSScriptRoot 'Get-FuelSurcharges.ps1'
if (-not (Test-Path $script)) { throw "Cannot find $script" }

# Validate the time before handing it to the scheduler.
$when = [datetime]::MinValue
if (-not [datetime]::TryParse($At, [ref]$when)) { throw "Could not read '$At' as a time of day. Try -At 07:15" }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $script) `
            -WorkingDirectory $PSScriptRoot

$trigger = New-ScheduledTaskTrigger -Daily -At $when

$settings = New-ScheduledTaskSettingsSet `
                -StartWhenAvailable `
                -DontStopIfGoingOnBatteries `
                -AllowStartIfOnBatteries `
                -ExecutionTimeLimit (New-TimeSpan -Minutes 15) `
                -MultipleInstances IgnoreNew

# Runs as you, only when you are logged on - no stored password, no elevation.
$principal = New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) `
                                        -LogonType Interactive -RunLevel Limited

if ($existing) { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false }

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
                       -Settings $settings -Principal $principal `
                       -Description 'Refreshes the ACE Courier fuel surcharge board from carrier websites.' | Out-Null

Write-Host "Registered '$TaskName' - runs daily at $($when.ToString('HH:mm'))." -ForegroundColor Green
Write-Host "  Script : $script" -ForegroundColor DarkGray
Write-Host "  Log    : $(Join-Path $PSScriptRoot 'data\run.log')" -ForegroundColor DarkGray
Write-Host "  Manage : taskschd.msc, or .\Register-FscTask.ps1 -Unregister" -ForegroundColor DarkGray

if ($RunNow) {
    Write-Host ''
    Write-Host 'Running once now...' -ForegroundColor Cyan
    Start-ScheduledTask -TaskName $TaskName
}
