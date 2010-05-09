;--------------------------------
;Include Modern UI
  !include "MUI.nsh"
;--------------------------------
!include Library.nsh

;Configuration
  ;General
Var STARTMENU_FOLDER
Var ALREADY_INSTALLED
!define PRODUCT "pop2owa"
!ifndef VERSION
	;Defaulf values if is used outside nant
	!define VERSION "vTest"
!endif
OutFile "${PRODUCT}_${VERSION}.exe"

Name "${PRODUCT} ${VERSION}"
InstallDir "$PROGRAMFILES\${PRODUCT}"
ShowInstDetails show
ShowUnInstDetails show
SetCompressor lzma ;bzip2 ;zlib
;SetCompressorDictSize 32
!packhdr tmpexe.tmp "UPX -9 -q --compress-icons=0 tmpexe.tmp"

;Get install folder from registry if available
InstallDirRegKey HKCU "Software\${PRODUCT}" ""

;--------------------------------
;Interface Settings

!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"
!define MUI_UNICON "${NSISDIR}\Contrib\Graphics\Icons\modern-uninstall.ico"
;!define MUI_STARTMENUPAGE_DEFAULTFOLDER ${PRODUCT}
;!define MUI_HEADERIMAGE
;!define MUI_HEADERIMAGE_BITMAP "logo.bmp"

;--------------------------------
;Language Selection Dialog Settings

;Remember the installer language
!define MUI_LANGDLL_REGISTRY_ROOT "HKCU"
!define MUI_LANGDLL_REGISTRY_KEY "Software\${PRODUCT}"
!define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"
;Remember the Start Menu Folder
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\${PRODUCT}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"
!define MUI_STARTMENUPAGE_DEFAULTFOLDER ${PRODUCT}
;--------------------------------
;Pages
; License page
!insertmacro MUI_PAGE_LICENSE "gpl.txt"
;Folder selection page
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_STARTMENU Application $STARTMENU_FOLDER
!insertmacro MUI_PAGE_INSTFILES
;Unistaller
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------

!define TEMP $R0


;--------------------------------
;Languages
  !insertmacro MUI_LANGUAGE "English"
  !insertmacro MUI_LANGUAGE "French"
  !insertmacro MUI_LANGUAGE "Spanish"
  !insertmacro MUI_LANGUAGE "Catalan"
  !insertmacro MUI_RESERVEFILE_LANGDLL

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
  ;!insertmacro DotNetSearch 3 5 "" "INSTALL_ABORT" ""
FunctionEnd

Section "Principal" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer


  IfFileExists "$INSTDIR\${PRODUCT}.exe" 0 new_installation
     StrCpy $ALREADY_INSTALLED 1
  new_installation:

  File ".\bin\Release\${PRODUCT}.exe"
;  File ".\bin\Release\config.xml"
  File ".\bin\Release\Microsoft.Exchange.WebServices.dll"
  File ".\bin\Release\NLog.dll"
  File ".\bin\Release\NLog.config"

;  File /oname=$INSTDIR\config.xml sample_config.xml

  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  createdirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\${PRODUCT}.lnk" "$INSTDIR\${PRODUCT}.exe"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Uninstall.lnk" "$INSTDIR\uninst.exe"
  !insertmacro MUI_STARTMENU_WRITE_END



SectionEnd


Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayName" "${PRODUCT} ${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayIcon" "$INSTDIR\makensis.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "URLInfoAbout" "http://www.pop2owa.com"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "Publisher" "Carlos Garcés"
SectionEnd

Section Uninstall

  Delete "$INSTDIR\${PRODUCT}.exe"
  Delete "$INSTDIR\uninst.exe"
  Delete "$INSTDIR\config.xml"
  Delete "$INSTDIR\Microsoft.Exchange.WebServices.dll"
  Delete "$INSTDIR\NLog.dll"
  Delete "$INSTDIR\NLog.config"

  RMDir "$INSTDIR\logs\logfile.txt"
  RMDir "$INSTDIR"
  
;  ;Remove shortcut
  !insertmacro MUI_STARTMENU_GETFOLDER Application ${TEMP}

  StrCmp ${TEMP} "" noshortcuts
         RMDir /R "$SMPROGRAMS\${TEMP}"
  noshortcuts:

  DeleteRegKey HKLM "Software\${PRODUCT}"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}"

;  SetAutoClose true
SectionEnd