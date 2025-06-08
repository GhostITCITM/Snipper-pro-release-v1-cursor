# Snipper Pro - Quick Setup Guide

## 🚀 One-Click Setup

### For First-Time Setup:
1. **Build**: Double-click `build.cmd`
2. **Register**: Double-click `run_as_admin.bat` 
3. **Launch**: Double-click `start_excel_with_snipper.bat`

That's it! The SNIPPER PRO tab should appear in Excel.

## 🔄 If You Need to Rebuild/Reinstall

### Complete Clean Rebuild:
```cmd
# 1. Clean old build files
Remove-Item -Recurse -Force SnipperCloneCleanFinal\bin, SnipperCloneCleanFinal\obj

# 2. Restore NuGet packages  
nuget.exe restore SnipperCloneCleanFinal.sln

# 3. Build
build.cmd

# 4. Re-register
run_as_admin.bat
```

### Just Reinstall Add-in:
```powershell
# Unregister old version
.\register_snipper_pro_simple.ps1 -Unregister

# Build latest
build.cmd

# Register new version  
.\register_snipper_pro_simple.ps1
```

## ✅ Verify Installation
```cmd
check_registration.ps1
verify_installation.ps1
```

## 🆘 Troubleshooting
- **Add-in not showing?** → Run `check_registration.ps1`
- **Build errors?** → Ensure .NET Framework 4.8 is installed
- **PDF not loading?** → Use `start_excel_with_snipper.bat` to launch Excel
- **Permission errors?** → Always run registration scripts as Administrator

---
**Need more details?** See the full `README.md` 