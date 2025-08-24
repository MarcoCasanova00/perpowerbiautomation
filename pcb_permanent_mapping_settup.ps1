# PC B - Permanent Drive Mapping Script
# Save as "Map-PowerBI-Drive.ps1" and run as Administrator

# Configuration
$DriveLetter = "P"  # Choose any available drive letter
$SharedPath = "\\PC-A-NAME\SharedPowerBI"  # Update with actual PC A name and share name
$Username = ""  # Leave empty to use current credentials
$Password = ""  # Leave empty to use current credentials

Write-Host "Setting up permanent drive mapping for Power BI files..."

try {
    # Method 1: Using net use command (most reliable)
    if ($Username -and $Password) {
        # With specific credentials
        $Result = cmd /c "net use ${DriveLetter}: `"$SharedPath`" /user:$Username $Password /persistent:yes 2>&1"
    } else {
        # With current user credentials
        $Result = cmd /c "net use ${DriveLetter}: `"$SharedPath`" /persistent:yes 2>&1"
    }
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Drive $DriveLetter`: mapped successfully to $SharedPath" -ForegroundColor Green
        Write-Host "✓ Mapping will persist after reboot" -ForegroundColor Green
        
        # Open the drive in Explorer
        Start-Process "explorer.exe" "${DriveLetter}:"
        Write-Host "✓ Opened drive in File Explorer" -ForegroundColor Green
        
    } else {
        Write-Host "❌ Error mapping drive: $Result" -ForegroundColor Red
        
        # Try alternative method
        Write-Host "Trying alternative method..."
        New-PSDrive -Name $DriveLetter -PSProvider FileSystem -Root $SharedPath -Persist -ErrorAction Stop
        Write-Host "✓ Drive mapped using PowerShell method" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure:" -ForegroundColor Yellow
    Write-Host "- PC A is turned on and accessible" -ForegroundColor Yellow
    Write-Host "- The shared folder exists and is accessible" -ForegroundColor Yellow
    Write-Host "- You have the correct permissions" -ForegroundColor Yellow
    Write-Host "- The drive letter $DriveLetter is not already in use" -ForegroundColor Yellow
}

# Verify the mapping
Write-Host "`nVerifying drive mapping..."
$Drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -eq $DriveLetter }
if ($Drives) {
    Write-Host "✓ Drive $DriveLetter`: is mapped to: $($Drives.DisplayRoot)" -ForegroundColor Green
} else {
    Write-Host "❌ Drive mapping verification failed" -ForegroundColor Red
}

# Show all network drives
Write-Host "`nCurrent network drives:"
Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -like "\\*" } | Format-Table Name, DisplayRoot, Used, Free -AutoSize

# Create shortcut on desktop (optional)
$CreateShortcut = Read-Host "`nCreate desktop shortcut to Power BI drive? (y/n)"
if ($CreateShortcut -eq 'y' -or $CreateShortcut -eq 'Y') {
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    $ShortcutPath = Join-Path $DesktopPath "PowerBI Files ($DriveLetter).lnk"
    
    $Shell = New-Object -ComObject WScript.Shell
    $Shortcut = $Shell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = "${DriveLetter}:"
    $Shortcut.Description = "Power BI Shared Files"
    $Shortcut.Save()
    
    Write-Host "✓ Desktop shortcut created: PowerBI Files ($DriveLetter)" -ForegroundColor Green
}

Read-Host "`nPress Enter to close"
