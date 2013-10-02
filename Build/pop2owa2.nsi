;--------------------------------
;Include Modern UI
  !include "MUI.nsh"
;--------------------------------
!include Library.nsh
!include x64.nsh

;Configuration
  ;General
Var STARTMENU_FOLDER
Var ALREADY_INSTALLED
!define PRODUCT "pop2owa"
!ifndef VERSION
	;Defaulf values if is used outside nant
	!define VERSION "0.0.0"
	OutFile "${PRODUCT}_${VERSION}.exe"
!endif

!define SOFWARE_UNINSTALL_KEY	"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}"

Name "${PRODUCT} ${VERSION}"
InstallDir "$PROGRAMFILES\${PRODUCT}"
ShowInstDetails show
ShowUnInstDetails show
SetCompressor lzma ;bzip2 ;zlib
;SetCompressorDictSize 32
;!packhdr tmpexe.tmp "upx -9 -q --compress-icons=0 tmpexe.tmp"

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
!insertmacro MUI_PAGE_LICENSE "..\temp\gpl.txt"
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
  
	VIProductVersion "${VERSION}.0"
	VIAddVersionKey /LANG=${LANG_ENGLISH} ProductName" "${PRODUCT}"
	VIAddVersionKey /LANG=${LANG_ENGLISH} "FileVersion" "${VERSION}"
	VIAddVersionKey /LANG=${LANG_ENGLISH} "InternalName" "${PRODUCT}.exe"
	;TODO Change description for different architectures?
	VIAddVersionKey /LANG=${LANG_ENGLISH} "FileDescription" "${PRODUCT} Installer"
	VIAddVersionKey /LANG=${LANG_ENGLISH} "LegalCopyright" ""

	
	
Function .onInit
  !insertmacro MUI_LANGDLL_DISPLAY
  ;!insertmacro DotNetSearch 3 5 "" "INSTALL_ABORT" ""
  	${If} ${RunningX64} 
		; disable registry redirection (enable access to 64-bit portion of registry)
		SetRegView 64
		; change install dir 
		StrCpy $INSTDIR "$PROGRAMFILES64\${PRODUCT}"
	${EndIf}
  
      ; Magic numbers from http://msdn.microsoft.com/en-us/library/ee942965.aspx
    ClearErrors
    ReadRegDWORD $0 HKLM "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" "Release"

    IfErrors NotDetected

    ${If} $0 >= 378389
        DetailPrint "Microsoft .NET Framework 4.5 is installed ($0)"
    ${Else}
    NotDetected:
        DetailPrint "Microsoft .NET Framework 4.5 NOT installed"
    ${EndIf}
  
FunctionEnd

Section "Principal" SEC01



  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer


  IfFileExists "$INSTDIR\${PRODUCT}.exe" 0 new_installation
     StrCpy $ALREADY_INSTALLED 1
  new_installation:

  File "..\temp\${PRODUCT}.exe"
;  File ".\bin\Release\config.xml"
  File "..\temp\Microsoft.Exchange.WebServices.dll"
  File "..\temp\NLog.dll"
  File "..\temp\NLog.config"

;  File /oname=$INSTDIR\config.xml sample_config.xml

  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  createdirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\${PRODUCT}.lnk" "$INSTDIR\${PRODUCT}.exe"
  CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Uninstall.lnk" "$INSTDIR\uninst.exe"
  !insertmacro MUI_STARTMENU_WRITE_END



SectionEnd


Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "DisplayName" "${PRODUCT} ${VERSION}"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "DisplayIcon" "$INSTDIR\makensis.exe"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "URLInfoAbout" "http://www.pop2owa.com"
  WriteRegStr HKLM "${SOFWARE_UNINSTALL_KEY}" "Publisher" "Carlos Garcés"
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
  DeleteRegKey HKLM "${SOFWARE_UNINSTALL_KEY}"

;  SetAutoClose true
SectionEnd