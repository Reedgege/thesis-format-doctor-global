; Thesis Format Doctor Global - Windows installer (NSIS)
; Wrap the Nuitka standalone folder into a single exe installer.
; Run at repo root: `makensis installer.nsi` (dist/thesis-format-doctor-global/ must exist)
; Version can be overridden: `makensis /DVERSION=2.0.0 installer.nsi`.
; ASCII-only on purpose: NSIS 3 reads the script as system ANSI (cp1252 on the
; English CI runner) unless a UTF-8 BOM is present; keeping the script ASCII
; avoids any mojibake without depending on the BOM.
Unicode true
!ifndef VERSION
  !define VERSION "dev"
!endif

!define APPNAME "Thesis Format Doctor Global"
!define APPDIR  "ThesisFormatDoctorGlobal"
!define EXE     "thesis-format-doctor-global.exe"
!define DIST    "dist\thesis-format-doctor-global"

Name "${APPNAME}"
OutFile "thesis-format-doctor-global-setup.exe"
InstallDir "$LOCALAPPDATA\Programs\${APPDIR}"
RequestExecutionLevel user          ; per-user install, no UAC prompt

!ifndef VERNUM
  !define VERNUM "0.0.0"
!endif
VIProductVersion "${VERNUM}.0"
VIAddVersionKey "ProductName" "${APPNAME}"
VIAddVersionKey "FileDescription" "${APPNAME} Installer"
VIAddVersionKey "FileVersion" "${VERNUM}"
VIAddVersionKey "ProductVersion" "${VERNUM}"
VIAddVersionKey "CompanyName" "ReedSkill"
VIAddVersionKey "LegalCopyright" "Copyright (c) ReedSkill"

!include "MUI2.nsh"
!define MUI_ABORTWARNING
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Section "Main" SecMain
  SetOutPath "$INSTDIR"
  File /r "${DIST}\*"

  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortCut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\${EXE}"
  CreateShortCut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\${EXE}"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "DisplayName" "${APPNAME}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "Publisher" "ReedSkill"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}" "URLInfoAbout" "https://reedskill.com"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\${APPNAME}.lnk"
  RMDir /r "$SMPROGRAMS\${APPNAME}"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPDIR}"
SectionEnd
