

## 1. Hourly Automation Script for PC A## 2. Permanent Drive Mapping for PC B## Setup Instructions:

### **PC A Setup (Run once):**
1. **Create the scripts folder:**
   ```powershell
   New-Item -ItemType Directory -Path "C:\Scripts" -Force
   ```

2. **Save the main script** as `C:\Scripts\Copy-PBIX-Hourly.ps1`

3. **Update the paths** in the script (your .pbix file location and shared folder)

4. **Run setup as Administrator:**
   ```powershell
   # Copy the setup portion and run it
   ```

### **PC B Setup (Run once):**
1. **Save the mapping script** and run as Administrator
2. **Update the shared path** with PC A's actual name

### **Auto-restart after PC A shutdown:**
The scheduled task will automatically restart when PC A boots up because:
- Task is set to "Start when available"
- Task has restart settings (3 attempts)
- Task runs every hour regardless of previous state

## Simple Commands for Quick Setup:

**PC A - Quick hourly task setup:**
```powershell
# Create hourly task (run as Admin)
schtasks /create /tn "PBIX-Copy" /tr "powershell.exe -File C:\Scripts\Copy-PBIX-Hourly.ps1" /sc hourly /ru %USERNAME%
```

**PC B - Quick permanent drive mapping:**
```cmd
net use P: \\PC-A-NAME\SharedFolder /persistent:yes
```

---------------------------------------------------------------------------------------------------------------------------------------

## Two different approaches:

### **Option 1: Fully Automated (Recommended)**
- The hourly script runs **automatically** every hour
- It copies whatever version of the .pbix file currently exists
- **You just need to save your Power BI file normally after refreshing**
- The script will pick up the updated file on the next hourly run

### **Option 2: Manual Trigger After Refresh**
If you want to copy immediately after refreshing, you'd need a **separate manual script**:

```powershell
# Quick manual copy script - Save as "Manual-Copy.ps1"
$SourceFile = "C:\Users\YourName\Documents\YourReport.pbix"
$SharedFolder = "C:\SharedPowerBI"

Copy-Item -Path $SourceFile -Destination $SharedFolder -Force
Write-Host "✓ File copied to shared folder immediately"
Read-Host "Press Enter to close"
```

## **Recommended Workflow:**
1. **Set up the hourly automation once** (using the first script)
2. **Work normally in Power BI:**
   - Refresh your data (Ctrl+R)
   - Save your file (Ctrl+S)
3. **That's it!** The automation handles the copying every hour

## **If you want immediate copying:**
Create both scripts:
- Keep the hourly automation running
- Use the manual script when you need immediate updates

