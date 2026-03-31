; ASCM Compliance Scanner - NSIS Installer
; Professional Windows installer for ASCM

!include "MUI2.nsh"
!include "x64.nsh"

; Define product information
!define PRODUCT_NAME "ASCM Compliance Scanner"
!define PRODUCT_VERSION "1.5.21"
!define PRODUCT_PUBLISHER "EUC Team"
!define PRODUCT_WEB_SITE "https://github.com/bhefe/ascm"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\Scan Software.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"

; Installer attributes
RequestExecutionLevel user
InstallDir ""  ; Will be set dynamically
BrandingText "ASCM Installer v${PRODUCT_VERSION}"
Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "ASCM-Installer-${PRODUCT_VERSION}.exe"
ShowInstDetails show
ShowUnInstDetails show

; MUI Settings
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_INSTFILES

; Language
!insertmacro MUI_LANGUAGE "English"

; Version Information
VIProductVersion "${PRODUCT_VERSION}.0"
VIAddVersionKey ProductName "${PRODUCT_NAME}"
VIAddVersionKey ProductVersion "${PRODUCT_VERSION}"
VIAddVersionKey CompanyName "${PRODUCT_PUBLISHER}"
VIAddVersionKey FileDescription "${PRODUCT_NAME} Installer"
VIAddVersionKey FileVersion "${PRODUCT_VERSION}.0"
VIAddVersionKey LegalCopyright "Copyright (c) 2026"
VIAddVersionKey OriginalFilename "ASCM-Installer-${PRODUCT_VERSION}.exe"

; Custom page for installation directory selection
Function .onInit
    ; Determine default installation directory
    ${If} ${RunningX64}
        StrCpy $INSTDIR "C:\Program Files\ASCM"
    ${Else}
        StrCpy $INSTDIR "C:\Program Files (x86)\ASCM"
    ${EndIf}
    
    ; Allow user to choose Desktop instead
    MessageBox MB_YESNO "Install ASCM to:$\n$\n• Yes: $INSTDIR$\n• No: Desktop" IDYES UseDefault IDNO UseDesktop
    
UseDesktop:
    StrCpy $INSTDIR "$DESKTOP\ascm"
    Goto DoneInit
    
UseDefault:
    ; Check if Program Files location needs elevation
    ${If} ${RunningX64}
        StrCpy $INSTDIR "C:\Program Files\ASCM"
    ${Else}
        StrCpy $INSTDIR "C:\Program Files (x86)\ASCM"
    ${EndIf}
    
DoneInit:
FunctionEnd

; Installation section
Section "Install ASCM"
    SetOutPath "$INSTDIR"
    
    ; Copy main executable
    SetDetailsPrint textonly
    DetailPrint "Copying Scan Software.exe..."
    SetDetailsPrint listonly
    SetOverwrite ifnewer
    File "dist\Scan Software.exe"
    
    ; Copy browser extension
    DetailPrint "Copying browser extension..."
    SetOverwrite ifnewer
    File /r "browser_extension\*.*"
    
    ; Copy supporting files
    DetailPrint "Copying supporting files..."
    SetOverwrite ifnewer
    File "README.md"
    File "requirements.txt"
    ${If} ${FileExists} "NOT_APPROVED.pdf"
        File "NOT_APPROVED.pdf"
    ${EndIf}
    
    ; Create uninstaller
    DetailPrint "Creating uninstaller..."
    WriteUninstaller "$INSTDIR\Uninstall-ASCM.exe"
    
    ; Create launcher batch file
    DetailPrint "Creating application launcher..."
    FileOpen $0 "$INSTDIR\Launch ASCM.bat" w
    FileWrite $0 "@echo off$\r$\ncd /d %~dp0$\r$\nstart Scan Software.exe$\r$\n"
    FileClose $0
    
    ; Create registry entries
    DetailPrint "Registering application..."
    WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\Scan Software.exe"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "DisplayName" "${PRODUCT_NAME}"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\Scan Software.exe"
    WriteRegStr HKLM "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\Uninstall-ASCM.exe"
    
    ; Create desktop shortcut
    DetailPrint "Creating desktop shortcut..."
    CreateDirectory "$SMPROGRAMS\ASCM"
    CreateShortCut "$DESKTOP\ASCM Compliance Scanner.lnk" "$INSTDIR\Scan Software.exe" "" "$INSTDIR\Scan Software.exe" 0 SW_SHOWNORMAL
    CreateShortCut "$SMPROGRAMS\ASCM\ASCM Compliance Scanner.lnk" "$INSTDIR\Scan Software.exe" "" "$INSTDIR\Scan Software.exe" 0 SW_SHOWNORMAL
    CreateShortCut "$SMPROGRAMS\ASCM\Uninstall ASCM.lnk" "$INSTDIR\Uninstall-ASCM.exe" "" "$INSTDIR\Uninstall-ASCM.exe" 0 SW_SHOWNORMAL
    
    SetDetailsPrint both
    DetailPrint "Installation complete!"
SectionEnd

; Uninstaller section
Section "Uninstall"
    SetDetailsPrint textonly
    DetailPrint "Removing ASCM files..."
    SetDetailsPrint listonly
    
    ; Remove files
    Delete "$INSTDIR\Scan Software.exe"
    Delete "$INSTDIR\Launch ASCM.bat"
    Delete "$INSTDIR\Uninstall-ASCM.exe"
    Delete "$INSTDIR\README.md"
    Delete "$INSTDIR\requirements.txt"
    Delete "$INSTDIR\NOT_APPROVED.pdf"
    
    ; Remove directories
    RMDir /r "$INSTDIR\browser_extension"
    RMDir "$INSTDIR"
    
    ; Remove shortcuts
    Delete "$DESKTOP\ASCM Compliance Scanner.lnk"
    RMDir /r "$SMPROGRAMS\ASCM"
    
    ; Remove registry entries
    DeleteRegKey HKLM "${PRODUCT_UNINST_KEY}"
    DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
    
    SetDetailsPrint both
    DetailPrint "Uninstallation complete!"
SectionEnd

; Finish Function
Function .onInstSuccess
    MessageBox MB_YESNO "Installation complete!$\n$\nWould you like to run ASCM now?" IDYES RunApp IDNO SkipRun
RunApp:
    ExecOpen "$INSTDIR\Scan Software.exe"
SkipRun:
FunctionEnd
