<#
Squeegee
   A PowerShell script by Ben Kennish (ben@kennish.net)

What's the best way to clean Windows?  With a Squeegee! ;)
Simply runs a lot of maintenance tasks to help ensure Windows is running well

#>

Set-StrictMode -Version Latest   # stricter rules = cleaner code  :)
$ErrorView = "DetailedView"

# this stuff is causing problems when a command errors
#$ErrorActionPreference = "Stop"
#$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8


$LASTEXITCODE = 0

# Create an object to hold the state of the disk drives' free space
# this is a bit messy as it exists outside the Write-Diskspace function but there
# are no static function variables in PowerShell
$diskState = New-Object -TypeName PSObject -Property @{
    spaceAtLastCheck = @{}
    spaceAtStart = @{}
}

function Write-DiskSpace
{
    param
    (
        [switch]$ShowTotals,
        [switch]$Reset
    )

    if ($Reset)
    {
        # pretend it's like the first time of running
         $diskState = New-Object -TypeName PSObject -Property @{
            spaceAtLastCheck = @{}
            spaceAtStart = @{}
        }
    }

    if ($ShowTotals)
    {
        # initialise this variable to add up
        $totalOverallChange = 0
    }

    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {
        # Only include local or removable drives
        ($_.DriveType -eq 2 -or $_.DriveType -eq 3) -and
        $_.VolumeName -notlike "*Google Drive*" -and
        $_.VolumeName -notlike "*Recovery*" -and
        $_.VolumeName -notlike "OEM"
    }

    $firstCall = $false
    if ($diskState.spaceAtLastCheck.Count -eq 0)
    {
        $firstCall = $true
    }


    foreach ($drive in $drives)
    {
        $driveLetter = $drive.DeviceID
        $freeSpaceMB = [math]::Round($drive.FreeSpace / 1MB)
        $label = if ($drive.VolumeName) { $drive.VolumeName } else { "No Label" }

        if ($firstCall)
        {
            # Initialize the hashtable for this drive
            $diskState.spaceAtLastCheck[$driveLetter] = $freeSpaceMB
            $diskState.spaceAtStart[$driveLetter] = $freeSpaceMB

            Write-Host "$("$driveLetter ($label)".PadRight(15)): $($freeSpaceMB.ToString("N0")) MB free" -NoNewline
        }
        else
        {
            $previousFreeSpace = $diskState.spaceAtLastCheck[$driveLetter]
            $spaceChange = $freeSpaceMB - $previousFreeSpace

            $diskState.spaceAtLastCheck[$driveLetter] = $freeSpaceMB
            #$diskState.spaceChangeSinceStart[$driveLetter] += $spaceChange  # tbh, we could just store the space at the start rather than constantly updating

            $changeString = if ($spaceChange -gt 0) { "+$($spaceChange.ToString("N0")) MB" } elseif ($spaceChange -lt 0) { "$($spaceChange.ToString("N0")) MB" } else { "0 MB" }

            Write-Host "$("$driveLetter ($label)".PadRight(15)): $($freeSpaceMB.ToString("N0")) MB free ($changeString)" -NoNewline
        }

        if ($ShowTotals)
        {
            $totalChangeThisDisk = $freeSpaceMB - $diskState.spaceAtStart[$driveLetter]
            $totalChangeAllDisks += $totalChangeThisDisk

            $totalChangeThisDiskHR = if ($totalChangeThisDisk -gt 0) { "+$($totalChangeThisDisk.ToString("N0")) MB" } elseif ($totalChangeThisDisk -lt 0) { "$($totalChangeThisDisk.ToString("N0")) MB" } else { "0 MB" }

            Write-Host "  TOTAL: $totalChangeThisDiskHR" -NoNewline
        }
        Write-Host ""
    }

    if ($ShowTotals)
    {
        $totalOverallChangeString = if ($totalChangeAllDisks -gt 0) { "+$($totalChangeAllDisks.ToString("N0")) MB" } elseif ($totalChangeAllDisks -lt 0) { "$($totalChangeAllDisks.ToString("N0")) MB" } else { "0 MB" }
        Write-Host "Space change overall: $totalOverallChangeString"
    }

}



# Check for Administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "You must run this script as an administrator." -ForegroundColor Red
    exit 1
}

Write-Host "You are running as an administrator.  Continuing..."

Write-Host "Current free disk space:"
Write-DiskSpace
Write-Host "------"

# Shut down all WSL instances and the lightweight VM
Write-Host "==== Shutting down the WSL lightweight VM..." -ForegroundColor Cyan
wsl --shutdown


Write-Host "==== Running chkdsk scan on all fixed disk drives..." -ForegroundColor Cyan
# Enumerate only fixed drives (DriveType 3)
Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object {
    $_.DriveType -eq 3 -and $_.VolumeName -notmatch 'Google Drive'
} | ForEach-Object {

    Write-Host ""
    Write-Host "======= Checking for filesystem errors on $($_.DeviceID) drive..." -ForegroundColor Green
    chkdsk /V /scan /perf $_.DeviceID

    if ($LASTEXITCODE -ne 0)
    {
        Write-Host "Filesystem errors found on $($_.DeviceID) drive. Exiting..." -ForegroundColor Red
        exit
    }
}


# Check for component store corruption and repair if necessary
Write-Host ""
Write-Host "==== Checking for component store corruption and repairing when necessary..." -ForegroundColor Cyan
dism /Online /Cleanup-Image /RestoreHealth
if ($LASTEXITCODE -ne 0)
{
    Write-Host "Command returned exit code $LASTEXITCODE" -ForegroundColor Red
    exit
}

Write-DiskSpace
Write-Host "------"

# Verify integrity of system files
Write-Host ""
Write-Host "==== Checking integrity of system files and replacing them if necessary..." -ForegroundColor Cyan
sfc /scannow
if ($LASTEXITCODE -ne 0)
{
    Write-Host "Command returned exit code $LASTEXITCODE" -ForegroundColor Red
    exit
}

Write-DiskSpace
Write-Host "------"

# Clean up the component store (WinSxS folder)
Write-Host ""
Write-Host "==== Cleaning up store of superseded components (WinSxS folder)..." -ForegroundColor Cyan
dism /Online /Cleanup-Image /StartComponentCleanup /ResetBase
if ($LASTEXITCODE -ne 0)
{
    Write-Host "Command returned exit code $LASTEXITCODE" -ForegroundColor Red
    exit
}

Write-DiskSpace
Write-Host "------"

# Run Disk Cleanup using preset #42
Write-Host ""
Write-Host "==== Running Disk Cleanup on all disk drives using preset #42..." -ForegroundColor Cyan
Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:42" -NoNewWindow -Wait

Write-DiskSpace
Write-Host "------"

# Run CCleaner's customized autoclean (ensure this path is correct)
Write-Host ""
Write-Host "==== Running CCleaner's customized autoclean..." -ForegroundColor Cyan
$cCleanerPath = "C:\Program Files\CCleaner\CCleaner64.exe"
if (Test-Path $cCleanerPath)
{
    #& $cCleanerPath "/AUTO"
    Start-Process -FilePath $cCleanerPath -ArgumentList "/AUTO" -NoNewWindow -Wait

    Write-DiskSpace
    Write-Host "------"
}
else
{
    Write-Host "CCleaner not found at $cCleanerPath. Skipping CCleaner autoclean step." -ForegroundColor Yellow
}


$localLowPath = [Environment]::GetFolderPath('LocalApplicationData') -replace '\\Local$', '\LocalLow'
$nvDxShaderCachePath = "$localLowPath\NVIDIA\PerDriverVersion\DXCache\"

Write-Host "NVIDIA DirectX shader cache: $nvDxShaderCachePath"

if (Test-Path $nvDxShaderCachePath)
{
    $answer = Read-Host "Do you want to clean it? (y/N) "
    if ($answer.ToLower() -eq "y")
    {
        $answer = ""

        do
        {
            Write-Host ""
            $handleOutput = & "handle.exe" $nvDxShaderCachePath 2>&1

            # Filter lines containing file locks and extract process names
            $lockedProcesses = $handleOutput |
            Select-String -Pattern "pid: (\d+).*" | # Match lines with process details
            ForEach-Object {
                # Extract process name
                if ($_ -match "^(.+?)\s+pid:")
                {
                    $matches[1].Trim() # Return the process name
                }
            } |
            Sort-Object -Unique

            # Output the list of locked processes
            if ($lockedProcesses.Count -gt 0)
            {
                Write-Host "Processes with locks in ${nvDxShaderCachePath}:"
                $lockedProcesses | ForEach-Object { Write-Host "- $_" }
            }
            else
            {
                Write-Host "No processes found with locks in $folderPath."
            }

            Write-Host ""
            $answer = Read-Host "(p)roceed, (r)efresh, or (Q)uit? "

        } while ($answer.ToLower() -eq "r")

        if ($answer.ToLower() -eq "p")
        {
            Write-Host ""
            Write-Host "==== Cleaning NVIDIA DirectX shader cache..." -ForegroundColor Cyan
            Remove-Item -Path "$nvDxShaderCachePath\*" -Force -Verbose -ErrorAction Continue

            Write-DiskSpace
            Write-Host "------"

        }
        else
        {
            Write-Host "Skipping cleaning of NVIDIA DirectX shader cache." -ForegroundColor Yellow
        }

    }
    else
    {
        Write-Host "Skipped cleaning of NVIDIA DirectX shader cache." -ForegroundColor Yellow
    }

}
else
{
    Write-Host "NVIDIA DirectX shader cache not found at $nvDxShaderCachePath. Skipping cleaning step." -ForegroundColor Yellow
}


Write-Host "==== Performing Docker system prune (unused images and anonymous volumes)..." -ForegroundColor Cyan
#$LASTEXITCODE = 0
#docker info

#if ($LASTEXITCODE -ne 0)
#{
#    Write-Host "Docker daemon is not running.  Skipping system prune." -ForegroundColor Yellow
#}
#else 
#{
    #Write-Host "Docker daemon is running.  Performing thorough system prune..." -ForegroundColor Cyan
    docker system prune --all --volumes

    Write-DiskSpace
    Write-Host "------"
#}

Write-Host ""
Write-Host "All done!"
Write-Host ""
Write-DiskSpace -ShowTotals
Write-Host "------"

Write-Host "Exiting in 2s..."
Start-Sleep -Seconds 2