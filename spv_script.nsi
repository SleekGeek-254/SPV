!include "MUI2.nsh"

!define APPNAME "SPV Installer"
!define COMPANY "SAMS"
!define INSTALL_DIR "$PROGRAMFILES\SPV"
!define VERSION "1.0"

Name "${COMPANY} ${APPNAME}"
OutFile "SPV_Installer_v${VERSION}.exe"
InstallDir "${INSTALL_DIR}"

; Request application privileges for Windows Vista and higher
RequestExecutionLevel admin

; Set the installer icon
Icon "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\icon.ico"

; Modern UI settings
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

Section "SPV Files" Section1
    SetOutPath "$INSTDIR"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\Secure PDF Viewer Example.adp"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\README.md"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\spvupdater.exe"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\icon.ico"
    
    SetOutPath "$INSTDIR\dist"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\dist\*.*"

    SetOutPath "$INSTDIR\example"
    File "D:\Company\1v1MeBro\EUTOPIA\ZOO\PYTHON\PDFVIEWER\001Dev\prod\example\*.*"
SectionEnd

Section "Create Desktop Shortcut" Section2
    SetShellVarContext all
    CreateShortCut "$DESKTOP\Secure PDF Viewer Example.lnk" "$INSTDIR\Secure PDF Viewer Example.adp" "" "$INSTDIR\icon.ico"
SectionEnd

Section "Create Start Menu Shortcut" Section3
    SetShellVarContext all
    CreateDirectory "$SMPROGRAMS\${COMPANY}"
    CreateShortCut "$SMPROGRAMS\${COMPANY}\Secure PDF Viewer Example.lnk" "$INSTDIR\Secure PDF Viewer Example.adp" "" "$INSTDIR\icon.ico"
    CreateShortCut "$SMPROGRAMS\${COMPANY}\SPV Updater.lnk" "$INSTDIR\spvupdater.exe" "" "$INSTDIR\icon.ico"
SectionEnd

; Uninstaller
Section "Uninstall"
    Delete "$INSTDIR\Secure PDF Viewer Example.adp"
    Delete "$INSTDIR\README.md"
    Delete "$INSTDIR\spvupdater.exe"
    Delete "$INSTDIR\icon.ico"
    Delete "$INSTDIR\dist\*.*"
    Delete "$INSTDIR\example\*.*"
    RMDir "$INSTDIR\dist"
    RMDir "$INSTDIR\example"
    RMDir "$INSTDIR"

    Delete "$DESKTOP\Secure PDF Viewer Example.lnk"

    Delete "$SMPROGRAMS\${COMPANY}\Secure PDF Viewer Example.lnk"
    Delete "$SMPROGRAMS\${COMPANY}\SPV Updater.lnk"
    RMDir "$SMPROGRAMS\${COMPANY}"

    DeleteRegKey HKLM "SOFTWARE\${COMPANY}\SPVINSTALL_DIR"
SectionEnd
