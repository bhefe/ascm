# ASCM Installer Options

This document explains the different ways to install ASCM Compliance Scanner.

## Option 1: Quick PowerShell Installer (No Admin Required)

**Best for:** Users who need ASCM quickly without NSIS setup

### Requirements
- Windows PowerShell 5.0+ (built-in on Windows 10/11)
- No admin rights required
- Built `Scan Software.exe` (run `python build_scan_exe.py` first)

### How to Use

1. **Run the installer:**
   ```powershell
   # Right-click on Install-ASCM.bat and select "Run as Administrator"
   # OR run in PowerShell:
   powershell -NoProfile -ExecutionPolicy Bypass -File Install-ASCM.ps1
   ```

2. **Choose installation location:**
   - Desktop (default, simplest)
   - Program Files (standard Windows location)
   - Custom path (enter your own)

3. **Installation complete!**
   - Exe copied to chosen location
   - Desktop shortcut created
   - Uninstall script provided

### What Gets Installed
- `Scan Software.exe` (the main executable)
- `browser_extension/` (Chrome/Edge extension files)
- `Launch ASCM.bat` (batch launcher)
- `Uninstall-ASCM.ps1` (removal script)
- Supporting files (README.md, requirements.txt, etc.)

## Option 2: Professional NSIS Installer (.exe)

**Best for:** Enterprise deployment, distribution to multiple users

### Requirements
- NSIS (Nullsoft Scriptable Install System)
- Windows 7+ (tested on 10/11)
- Built `Scan Software.exe`

### Installation Steps

#### 1. Install NSIS (One-time setup)
```powershell
# Option A: Using Chocolatey (requires admin)
choco install nsis

# Option B: Download from website
# Visit: https://nsis.sourceforge.io/Download
# Download the installer exe and run it
```

#### 2. Build the NSIS Installer
```powershell
# Run the build script
& ".\build_nsis_installer.ps1"

# Follow the prompts:
# - Confirm all files are present
# - Enter version (default: 1.5.21)
# - Wait for build to complete

# Output: ASCM-Installer-1.5.21.exe (in project root)
```

#### 3. Distribute & Run
- Share `ASCM-Installer-1.5.21.exe` with users
- Users run the installer
- Select installation directory:
  - Program Files\ASCM (recommended for enterprise)
  - Desktop\ascm (simplest for single user)
- Installer handles:
  - File copying
  - Registry entries
  - Shortcut creation
  - Uninstaller setup

### Installing from NSIS Installer

**As end user:**
1. Download `ASCM-Installer-1.5.21.exe`
2. Double-click to run
3. Choose installation location:
   - **Yes** → Install to Program Files (or C:\Program Files\ASCM)
   - **No** → Install to Desktop\ascm
4. Click Install
5. Finish dialog offers to run ASCM immediately
6. Done! Desktop shortcut created

**Uninstalling:**
- Use Windows Settings → Apps → ASCM Compliance Scanner → Uninstall
- OR run: `ASCM-Installer-1.5.21.exe` again
- OR run: Control Panel → Programs → Uninstall a program → ASCM

## Option 3: Manual Installation

**Best for:** Developers, advanced users

No installer needed - copy files manually:

```powershell
# Create directory
New-Item -Type Directory "C:\Program Files\ASCM" -Force

# Copy files
Copy-Item "dist\Scan Software.exe" "C:\Program Files\ASCM\"
Copy-Item "browser_extension" "C:\Program Files\ASCM\" -Recurse

# Run
& "C:\Program Files\ASCM\Scan Software.exe"
```

## Comparison Table

| Feature | PowerShell Installer | NSIS Installer | Manual |
|---------|----------------------|----------------|--------|
| User Experience | Good (dialog UI) | Excellent (wizard) | Poor (technical) |
| Setup Required | None | NSIS install needed | Command line only |
| Admin Rights | Not required | Not required | Optional |
| Registry Entries | No | Yes (clean uninstall) | Manual if needed |
| Shortcuts | Desktop | Desktop + Start Menu | Manual |
| Enterprise Ready | Fair | Excellent | No |
| Setup Time | 1-2 minutes | 5-10 minutes (first time) | 5 minutes |
| File Size | ~40 MB (script only) | ~33 MB (exe) | ~33 MB (exe) |

## Troubleshooting

### PowerShell Installer Issues

**"Install-ASCM.ps1 cannot be loaded"**
```powershell
# Run this first to enable script execution:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**"Scan Software.exe not found"**
- Ensure you've built the exe: `python build_scan_exe.py`
- Check the build completed: Look in `dist/` folder

**Installation fails silently**
- Run with elevated privileges (Run as Administrator)
- Check Windows Defender doesn't block the exe

### NSIS Installer Issues

**"NSIS not found"**
- Download from https://nsis.sourceforge.io/Download
- Install to: `C:\Program Files\NSIS`
- Then run `build_nsis_installer.ps1` again

**"makensis.exe not found"**
- NSIS installed but not in standard location
- Edit `build_nsis_installer.ps1` to update the path

**Installer won't build**
- Check all files exist (see "What Gets Installed" section)
- Run PowerShell as Administrator
- Ensure `dist\Scan Software.exe` is built and not in use

## Next Steps

### For Quick Testing
1. Run PowerShell installer: `Install-ASCM.ps1`
2. Choose Desktop location
3. Double-click desktop shortcut to test

### For Enterprise Deployment
1. Install NSIS on your build machine
2. Run `build_nsis_installer.ps1`
3. Test the resulting `.exe` installer
4. Distribute `ASCM-Installer-1.5.21.exe` to users
5. Users run installer themselves

### Browser Extension Setup
After installation (either method):

**Chrome:**
1. Open `chrome://extensions/`
2. Enable "Developer mode" (top right)
3. Click "Load unpacked"
4. Navigate to: `[Installation Path]\browser_extension`
5. Done! Extension ready to use

**Edge:**
1. Open `edge://extensions/`
2. Enable "Developer mode" (bottom left)
3. Click "Load unpacked"
4. Navigate to: `[Installation Path]\browser_extension`
5. Done! Extension ready to use

## File Structure After Installation

### Desktop Installation
```
Desktop/
├── ascm/
│   ├── Scan Software.exe
│   ├── Launch ASCM.bat
│   ├── Uninstall-ASCM.ps1
│   ├── browser_extension/
│   ├── README.md
│   ├── requirements.txt
│   └── NOT_APPROVED.pdf
└── ASCM Compliance Scanner.lnk (shortcut)
```

### Program Files Installation
```
Program Files/ASCM/
├── Scan Software.exe
├── Launch ASCM.bat
├── Uninstall-ASCM.exe
├── browser_extension/
├── README.md
├── requirements.txt
└── NOT_APPROVED.pdf
```

## Creating Custom Installers

### Edit PowerShell Installer
Edit `Install-ASCM.ps1`:
- Change `filesToCopy` array to include/exclude files
- Modify shortcut names
- Add custom registry entries in `Install-ASCM` function

### Edit NSIS Installer
Edit `installer/ascm.nsi`:
- Modify `PRODUCT_*` variables for custom branding
- Change installation directories
- Add/remove shortcuts or registry keys
- Customize welcome/finish messages

### Build Custom NSIS
```powershell
# Modify the .nsi file, then rebuild:
& ".\build_nsis_installer.ps1"
```

## Version History

- **1.5.21** - Browser extension fix, CUDA/NVIDIA tools, graceful error handling
- **1.5.20** - Error handling for permission issues
- **1.5.0+** - Email hyperlinks, software consolidation
- **1.3.0** - XLSX report generation

## Support

- **GitHub:** https://github.com/bhefe/ascm
- **Issues:** Report problems via GitHub Issues
- **Questions:** Check README.md and INSTALLATIONS.md

---

**Last Updated:** March 31, 2026
**Version:** 1.5.21
