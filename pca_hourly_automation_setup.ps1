# STEP 1: Main Copy Script - Save as "Copy-PBIX-Hourly.ps1"
# This is the script that runs every hour

param(
    [string]$SourceFile = "C:\Users\YourName\Documents\YourReport.pbix",
    [string]$SharedFolder = "C:\SharedPowerBI"  # Local path on PC A
)

$LogFile = "C:\Logs\PBIX-Copy.log"

# Create log directory if needed
$LogDir = Split-Path $LogFile -Parent
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force }

function Write-Log {
    param([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Tee-Object -FilePath $LogFile -Append
}

try {
    Write-Log "=== Hourly PBIX Copy Started ==="
    
    # Check if source file exists
    if (!(Test-Path $SourceFile)) {
        Write-Log "ERROR: Source file not found: $SourceFile"
        exit 1
    }
    
    # Check if shared folder exists
    if (!(Test-Path $SharedFolder)) {
        Write-Log "ERROR: Shared folder not found: $SharedFolder"
        exit 1
    }
    
    # Get file info
    $FileInfo = Get-Item $SourceFile
    $FileName = $FileInfo.Name
    $FileSize = [math]::Round($FileInfo.Length / 1MB, 2)
    
    Write-Log "Copying $FileName ($FileSize MB) to shared folder..."
    
    # Copy the file
    Copy-Item -Path $SourceFile -Destination $SharedFolder -Force
    
    # Verify copy
    $DestFile = Join-Path $SharedFolder $FileName
    if (Test-Path $DestFile) {
        $DestSize = [math]::Round((Get-Item $DestFile).Length / 1MB, 2)
        if ($DestSize -eq $FileSize) {
            Write-Log "SUCCESS: File copied successfully ($DestSize MB)"
        } else {
            Write-Log "WARNING: File size mismatch (Source: $FileSize MB, Dest: $DestSize MB)"
        }
    } else {
        Write-Log "ERROR: File not found in destination after copy"
    }
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
} finally {
    Write-Log "=== Hourly PBIX Copy Completed ==="
}

# STEP 2: Task Scheduler Setup Script - Save as "Setup-Hourly-Task.ps1"
# Run this ONCE as Administrator to set up the hourly task

$TaskName = "PBIX-Hourly-Copy"
$ScriptPath = "C:\Scripts\Copy-PBIX-Hourly.ps1"  # Update this path
$LogPath = "C:\Logs\PBIX-Task.log"

Write-Host "Setting up hourly task: $TaskName"

try {
    # Create the task action
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`""
    
    # Create hourly trigger
    $Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
    
    # Task settings
    $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    
    # Run as current user
    $Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
    
    # Register the task
    Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description "Hourly copy of PBIX file to shared folder" -Force
    
    Write-Host "✓ Task created successfully!"
    Write-Host "✓ Will run every hour starting now"
    Write-Host "✓ Auto-restarts on failure"
    Write-Host "✓ Runs even on battery power"
    
    # Start the task immediately
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "✓ Task started immediately"
    
} catch {
    Write-Host "❌ Error creating task: $($_.Exception.Message)" -ForegroundColor Red
}

# STEP 3: Startup Script for Auto-Restart - Save as "Startup-Task-Check.ps1"
# Add this to Windows Startup folder or create another scheduled task

$TaskName = "PBIX-Hourly-Copy"

try {
    $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    $TaskState = $Task.State
    
    if ($TaskState -ne "Running" -and $TaskState -ne "Ready") {
        Write-Host "Task is $TaskState, attempting to restart..."
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "✓ Task restarted"
    } else {
        Write-Host "✓ Task is running normally ($TaskState)"
    }
} catch {
    Write-Host "❌ Task not found or error: $($_.Exception.Message)"
}
