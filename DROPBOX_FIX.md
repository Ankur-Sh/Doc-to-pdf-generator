# Fixing Permission Errors in Dropbox Folders

## Problem

If you're getting "Permission denied" errors and your project is in a Dropbox folder, this is likely a Dropbox sync issue.

## Quick Fixes

### Option 1: Wait for Dropbox Sync (Easiest)

1. Check Dropbox icon in system tray
2. Wait until it shows "Up to date" (no sync in progress)
3. Run the script again

### Option 2: Close Microsoft Word

1. Press `Ctrl + Shift + Esc` to open Task Manager
2. Look for "WINWORD.EXE" processes
3. Right-click → End Task
4. Run the script again

### Option 3: Move Project Outside Dropbox (Recommended)

**Best solution for avoiding sync conflicts:**

1. Copy the entire project folder to a local location:
   - `C:\Users\YourName\Documents\convert text to pdf`
   - Or `C:\Projects\convert text to pdf`

2. Run the script from the new location

3. **Why this helps:**
   - Dropbox syncs files in the background
   - This can lock files temporarily
   - Local folders don't have this issue

### Option 4: Pause Dropbox Temporarily

1. Right-click Dropbox icon in system tray
2. Click "Pause syncing" → "2 hours"
3. Run the script
4. Resume Dropbox when done

### Option 5: Exclude Output Folders from Dropbox

1. Right-click Dropbox folder → Dropbox → Preferences
2. Go to "Sync" tab
3. Add `output` and `output-docx` folders to exclusions
4. This prevents Dropbox from syncing output files

## Prevention

**Best Practice:** Keep your project in a local folder (not Dropbox), and only sync the source files (`files_to_convert` folder) if needed.

## Still Having Issues?

1. **Check file permissions:**
   - Right-click project folder → Properties → Security
   - Make sure your user has "Full control"

2. **Run as Administrator:**
   - Right-click Command Prompt → Run as administrator
   - Navigate to project folder
   - Run the script

3. **Check antivirus:**
   - Temporarily disable antivirus
   - Or add project folder to exclusions

4. **Delete locked files manually:**
   ```cmd
   del /F /Q "output-docx\*.docx"
   del /F /Q "output-docx\*.pdf"
   ```







