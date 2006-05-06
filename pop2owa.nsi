;--------------------------------
;Include Modern UI
  !include "MUI.nsh"
;--------------------------------
;Configuration
  ;General
Var STARTMENU_FOLDER
!define PRODUCT "pop2owa"
!define VERSION "v1.0RC"
Name "${PRODUCT} ${VERSION}"
OutFile "${PRODUCT}_${VERSION}.exe"
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
  !insertmacro MUI_LANGUAGE "Catalan"
  !insertmacro MUI_LANGUAGE "English"
  !insertmacro MUI_LANGUAGE "French"
  !insertmacro MUI_LANGUAGE "Spanish"
  !insertmacro MUI_RESERVEFILE_LANGDLL

Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
FunctionEnd

Section "Principal" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "${PRODUCT}.exe"

;  WriteIniStr "$INSTDIR\VBDOXAddin.url" "InternetShortcut" "URL" "http://vbdoxaddin.carlos-garces.com"

  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  createdirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\${PRODUCT}.lnk" "$INSTDIR\${PRODUCT}.exe"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Uninstall.lnk" "$INSTDIR\uninst.exe"
  !insertmacro MUI_STARTMENU_WRITE_END

  ;Test the last value inserted (v0.8)
  ReadRegStr $R0 HKCU "Software\${PRODUCT}" "FormBasedAuth"
  StrCmp $R0 "" write done

write:
; Write defalut aplication values
  WriteRegStr HKCU "Software\${PRODUCT}" "ExchangeServer" "http://mail.yourcompany.com"
  WriteRegStr HKCU "Software\${PRODUCT}" "IP" "127.0.0.1"
  WriteRegDWORD HKCU "Software\${PRODUCT}" "POP3" 0x6E
  WriteRegDWORD HKCU "Software\${PRODUCT}" "SMTP" 0x19
  WriteRegDWORD HKCU "Software\${PRODUCT}" "Saveinsent" 0x1
  WriteRegDWORD HKCU "Software\${PRODUCT}" "FormBasedAuth" 0x1

done:

SectionEnd


Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayName" "${PRODUCT} ${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayIcon" "$INSTDIR\makensis.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "URLInfoAbout" "http://pop2owa.sourceforge.net"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "Publisher" "Carlos Garcés"
SectionEnd


;Function un.onUninstSuccess
;  HideWindow
;  MessageBox MB_ICONINFORMATION|MB_OK "La desinstalación de ${PRODUCT} ${VERSION} finalizó satisfactoriamente."
;FunctionEnd

;Function un.onInit
;  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "¿Está completamente seguro que desea desinstalar ${PRODUCT} ${VERSION} junto con todos sus componentes?" IDYES +2
;  Abort
;FunctionEnd

Section Uninstall

  Delete "$INSTDIR\${PRODUCT}.exe"
  Delete "$INSTDIR\uninst.exe"

;  ;Remove shortcut
  !insertmacro MUI_STARTMENU_GETFOLDER Application ${TEMP}

  StrCmp ${TEMP} "" noshortcuts
         ;Delete "$SMPROGRAMS\${TEMP}\${PRODUCT}.lnk"
         ;Delete "$SMPROGRAMS\${TEMP}\Uninstall.lnk"
         RMDir /R "$SMPROGRAMS\${TEMP}"
  noshortcuts:
  RMDir "$INSTDIR"

  DeleteRegKey HKCU "Software\${PRODUCT}"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}"

;  SetAutoClose true
SectionEnd



