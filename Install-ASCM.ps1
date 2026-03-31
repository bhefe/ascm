# ASCM Compliance Scanner - Installation Script
# This script installs ASCM to a user-selected location and creates shortcuts

param(
    [string]$InstallPath
)

# Function to display the UI
function Show-InstallationDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ASCM Installer"
    $form.Width = 500
    $form.Height = 280
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    # Title
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select Installation Location"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    # Description
    $desc = New-Object System.Windows.Forms.Label
    $desc.Text = "Choose where to install ASCM Compliance Scanner"
    $desc.AutoSize = $true
    $desc.Location = New-Object System.Drawing.Point(20, 50)
    $desc.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($desc)

    # Desktop button
    $btnDesktop = New-Object System.Windows.Forms.Button
    $btnDesktop.Text = "Desktop"
    $btnDesktop.Width = 150
    $btnDesktop.Height = 40
    $btnDesktop.Location = New-Object System.Drawing.Point(20, 100)
    $btnDesktop.BackColor = [System.Drawing.Color]::LightBlue
    $btnDesktop.Add_Click({
        $script:SelectedPath = "$env:USERPROFILE\Desktop\ascm"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnDesktop)

    # Program Files button
    $btnProgFiles = New-Object System.Windows.Forms.Button
    $btnProgFiles.Text = "Program Files"
    $btnProgFiles.Width = 150
    $btnProgFiles.Height = 40
    $btnProgFiles.Location = New-Object System.Drawing.Point(200, 100)
    $btnProgFiles.BackColor = [System.Drawing.Color]::LightGreen
    $btnProgFiles.Add_Click({
        $script:SelectedPath = "C:\Program Files\ASCM"
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnProgFiles)

    # Custom path option
    $lblCustom = New-Object System.Windows.Forms.Label
    $lblCustom.Text = "Or Enter Custom Path:"
    $lblCustom.AutoSize = $true
    $lblCustom.Location = New-Object System.Drawing.Point(20, 160)
    $form.Controls.Add($lblCustom)

    $textboxCustom = New-Object System.Windows.Forms.TextBox
    $textboxCustom.Width = 380
    $textboxCustom.Location = New-Object System.Drawing.Point(20, 185)
    $textboxCustom.Text = "$env:USERPROFILE\Desktop\ascm"
    $form.Controls.Add($textboxCustom)

    # Install button
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = "Install"
    $btnInstall.Width = 100
    $btnInstall.Height = 35
    $btnInstall.Location = New-Object System.Drawing.Point(300, 220)
    $btnInstall.BackColor = [System.Drawing.Color]::LightGreen
    $btnInstall.Add_Click({
        if ($textboxCustom.Text) {
            $script:SelectedPath = $textboxCustom.Text
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnInstall)

    # Cancel button
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Width = 100
    $btnCancel.Height = 35
    $btnCancel.Location = New-Object System.Drawing.Point(410, 220)
    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $result = $form.ShowDialog()
    return $result
}

# Main installation logic
function Install-ASCM {
    param([string]$Path)

    Write-Host "`n🔒 ASCM Compliance Scanner - Installation`n" -ForegroundColor Cyan
    Write-Host "Installing to: $Path`n" -ForegroundColor Yellow

    # Create installation directory
    if (-not (Test-Path $Path)) {
        Write-Host "Creating directory..." -ForegroundColor Gray
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    # Get the source exe
    $sourceExe = "c:\Users\EUC User\Desktop\ascm\dist\Scan Software.exe"
    
    if (-not (Test-Path $sourceExe)) {
        Write-Host "❌ Error: Scan Software.exe not found at $sourceExe" -ForegroundColor Red
        Write-Host "Please ensure the exe has been built. Run: python build_scan_exe.py" -ForegroundColor Yellow
        return $false
    }

    # Copy exe
    Write-Host "Copying Scan Software.exe..." -ForegroundColor Gray
    Copy-Item -Path $sourceExe -Destination "$Path\Scan Software.exe" -Force
    
    # Copy browser extension
    $extensionSource = "c:\Users\EUC User\Desktop\ascm\browser_extension"
    if (Test-Path $extensionSource) {
        Write-Host "Copying browser extension..." -ForegroundColor Gray
        Copy-Item -Path $extensionSource -Destination "$Path\browser_extension" -Recurse -Force
    }

    # Copy other important files
    $filesToCopy = @(
        "requirements.txt",
        "README.md",
        "NOT_APPROVED.pdf"
    )
    
    foreach ($file in $filesToCopy) {
        $sourceFile = "c:\Users\EUC User\Desktop\ascm\$file"
        if (Test-Path $sourceFile) {
            Write-Host "Copying $file..." -ForegroundColor Gray
            Copy-Item -Path $sourceFile -Destination "$Path\$file" -Force
        }
    }

    # Create uninstaller
    Write-Host "Creating uninstaller..." -ForegroundColor Gray
    $uninstallerScript = @"
# ASCM Uninstaller
`$path = Split-Path -Parent `$MyInvocation.MyCommand.Path
if ((Read-Host "Remove ASCM from ``$path`` ? (yes/no)") -eq "yes") {
    Remove-Item -Path `$path -Recurse -Force
    Write-Host "ASCM uninstalled successfully." -ForegroundColor Green
}
"@
    
    Set-Content -Path "$Path\Uninstall-ASCM.ps1" -Value $uninstallerScript -Encoding UTF8

    # Create launcher batch file
    Write-Host "Creating launcher..." -ForegroundColor Gray
    $launcherBatch = @"
@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process 'Scan Software.exe' -NoNewWindow"
"@
    
    Set-Content -Path "$Path\Launch ASCM.bat" -Value $launcherBatch -Encoding ASCII

    # Create desktop shortcut
    $desktopPath = "$env:USERPROFILE\Desktop"
    Write-Host "Creating desktop shortcut..." -ForegroundColor Gray
    
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut("$desktopPath\ASCM Compliance Scanner.lnk")
    $shortcut.TargetPath = "$Path\Scan Software.exe"
    $shortcut.WorkingDirectory = $Path
    $shortcut.Description = "ASCM Compliance Scanner - Scan installed software against approval lists"
    $shortcut.Save()

    Write-Host "`n✅ Installation Complete!`n" -ForegroundColor Green
    Write-Host "Installation Details:" -ForegroundColor Cyan
    Write-Host "  Location: $Path" -ForegroundColor White
    Write-Host "  Shortcut: $desktopPath\ASCM Compliance Scanner.lnk" -ForegroundColor White
    Write-Host "  Launcher: $Path\Launch ASCM.bat" -ForegroundColor White
    Write-Host "  Uninstall: $Path\Uninstall-ASCM.ps1" -ForegroundColor White
    Write-Host "`nTo run ASCM:" -ForegroundColor Cyan
    Write-Host "  - Double-click the desktop shortcut, OR" -ForegroundColor White
    Write-Host "  - Run: $Path\Scan Software.exe" -ForegroundColor White
    Write-Host "  - Run: $Path\Launch ASCM.bat" -ForegroundColor White
    Write-Host "  - Run: & '$Path\Scan Software.exe'" -ForegroundColor White
    Write-Host "`nBrowser Extension:" -ForegroundColor Cyan
    Write-Host "  Load the extension from: $Path\browser_extension" -ForegroundColor White
    Write-Host "  Chrome: chrome://extensions/ → Developer mode → Load unpacked" -ForegroundColor White
    Write-Host "  Edge: edge://extensions/ → Developer mode → Load unpacked`n" -ForegroundColor White

    return $true
}

# Main execution
$dialogResult = Show-InstallationDialog

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    if (Install-ASCM -Path $SelectedPath) {
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } else {
        Write-Host "Installation failed. Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
} else {
    Write-Host "Installation cancelled." -ForegroundColor Yellow
}
