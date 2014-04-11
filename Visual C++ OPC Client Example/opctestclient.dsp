# Microsoft Developer Studio Project File - Name="opctestclient" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=opctestclient - Win32 Unicode Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "opctestclient.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "opctestclient.mak" CFG="opctestclient - Win32 Unicode Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "opctestclient - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "opctestclient - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE "opctestclient - Win32 Unicode Debug" (based on "Win32 (x86) Application")
!MESSAGE "opctestclient - Win32 Unicode Release" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""$/src32/KEPServerEx/opctestclient distributable source code version", RTQBAAAA"
# PROP Scc_LocalPath "."
CPP=cl.exe
F90=df.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "opctestclient - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_WIN32_DCOM" /D "_CLIENT" /YX"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 /nologo /subsystem:windows /machine:I386 /out:"Release\opcquickclient.exe"

!ELSEIF  "$(CFG)" == "opctestclient - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD F90 /browser
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_WIN32_DCOM" /D "_CLIENT" /FR /YX"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /debug /machine:I386 /out:"Debug\opcquickclient.exe" /pdbtype:sept
# SUBTRACT LINK32 /nodefaultlib

!ELSEIF  "$(CFG)" == "opctestclient - Win32 Unicode Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "opctestclient___Win32_Unicode_Debug"
# PROP BASE Intermediate_Dir "opctestclient___Win32_Unicode_Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "debugu"
# PROP Intermediate_Dir "debugu"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /I "..\common" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_WIN32_DCOM" /D "_CLIENT" /YX"stdafx.h" /FD /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_WIN32_DCOM" /D "_CLIENT" /D "_UNICODE" /YX"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /out:"..\opcquickclient.exe" /pdbtype:sept
# SUBTRACT BASE LINK32 /nodefaultlib
# ADD LINK32 /nologo /entry:"wWinMainCRTStartup" /subsystem:windows /debug /machine:I386 /out:"Debugu\opcquickclient.exe" /pdbtype:sept
# SUBTRACT LINK32 /nodefaultlib

!ELSEIF  "$(CFG)" == "opctestclient - Win32 Unicode Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "opctestclient___Win32_Unicode_Release"
# PROP BASE Intermediate_Dir "opctestclient___Win32_Unicode_Release"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "releaseu"
# PROP Intermediate_Dir "releaseu"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_CLIENT" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_WIN32_DCOM" /D "_CLIENT" /D "_UNICODE" /YX"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386 /out:"..\opcquickclient.exe"
# ADD LINK32 /nologo /entry:"wWinMainCRTStartup" /subsystem:windows /machine:I386 /out:"Releaseu\opcquickclient.exe"

!ENDIF 

# Begin Target

# Name "opctestclient - Win32 Release"
# Name "opctestclient - Win32 Debug"
# Name "opctestclient - Win32 Unicode Debug"
# Name "opctestclient - Win32 Unicode Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\advisesink.cpp
# End Source File
# Begin Source File

SOURCE=.\datasink20.cpp
# End Source File
# Begin Source File

SOURCE=.\document.cpp
# End Source File
# Begin Source File

SOURCE=.\editfilters.cpp
# End Source File
# Begin Source File

SOURCE=.\eventview.cpp
# End Source File
# Begin Source File

SOURCE=.\globals.cpp
# End Source File
# Begin Source File

SOURCE=.\group.cpp
# End Source File
# Begin Source File

SOURCE=.\grouppropertysheet.cpp
# End Source File
# Begin Source File

SOURCE=.\groupview.cpp
# End Source File
# Begin Source File

SOURCE=.\imagebutton.cpp
# End Source File
# Begin Source File

SOURCE=.\item.cpp
# End Source File
# Begin Source File

SOURCE=.\itemadddlg.cpp
# End Source File
# Begin Source File

SOURCE=.\itempropertiesdlg.cpp
# End Source File
# Begin Source File

SOURCE=.\itemview.cpp
# End Source File
# Begin Source File

SOURCE=.\itemwritedlg.cpp
# End Source File
# Begin Source File

SOURCE=.\listeditctrl.cpp
# End Source File
# Begin Source File

SOURCE=.\mainwnd.cpp
# End Source File
# Begin Source File

SOURCE=.\opccomn_i.c
# End Source File
# Begin Source File

SOURCE=.\opcda_i.c
# End Source File
# Begin Source File

SOURCE=.\OpcEnum_i.c
# End Source File
# Begin Source File

SOURCE=.\opctestclient.cpp
# End Source File
# Begin Source File

SOURCE=.\safearray.cpp
# End Source File
# Begin Source File

SOURCE=.\server.cpp
# End Source File
# Begin Source File

SOURCE=.\serverenumgroupsdlg.cpp
# End Source File
# Begin Source File

SOURCE=.\servergeterrorstringdlg.cpp
# End Source File
# Begin Source File

SOURCE=.\servergroupbynamedlg.cpp
# End Source File
# Begin Source File

SOURCE=.\serverpropertysheet.cpp
# End Source File
# Begin Source File

SOURCE=.\shutdownsink.cpp
# End Source File
# Begin Source File

SOURCE=.\smarttooltip.cpp
# End Source File
# Begin Source File

SOURCE=.\statusbartext.cpp
# End Source File
# Begin Source File

SOURCE=.\timestmp.cpp
# End Source File
# Begin Source File

SOURCE=.\updateintervaldlg.cpp
# End Source File
# Begin Source File

SOURCE=.\versioninfo.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\advisesink.h
# End Source File
# Begin Source File

SOURCE=.\datasink20.h
# End Source File
# Begin Source File

SOURCE=.\document.h
# End Source File
# Begin Source File

SOURCE=.\editfilters.h
# End Source File
# Begin Source File

SOURCE=.\eventview.h
# End Source File
# Begin Source File

SOURCE=.\fixedsharedfile.h
# End Source File
# Begin Source File

SOURCE=.\globals.h
# End Source File
# Begin Source File

SOURCE=.\group.h
# End Source File
# Begin Source File

SOURCE=.\grouppropertysheet.h
# End Source File
# Begin Source File

SOURCE=.\groupview.h
# End Source File
# Begin Source File

SOURCE=.\imagebutton.h
# End Source File
# Begin Source File

SOURCE=.\item.h
# End Source File
# Begin Source File

SOURCE=.\itemadddlg.h
# End Source File
# Begin Source File

SOURCE=.\itempropertiesdlg.h
# End Source File
# Begin Source File

SOURCE=.\itemview.h
# End Source File
# Begin Source File

SOURCE=.\itemwritedlg.h
# End Source File
# Begin Source File

SOURCE=.\listeditctrl.h
# End Source File
# Begin Source File

SOURCE=.\mainwnd.h
# End Source File
# Begin Source File

SOURCE=.\opccomn.h
# End Source File
# Begin Source File

SOURCE=.\opcda.h
# End Source File
# Begin Source File

SOURCE=.\OpcEnum.h
# End Source File
# Begin Source File

SOURCE=.\opcerrors.h
# End Source File
# Begin Source File

SOURCE=.\opcprops.h
# End Source File
# Begin Source File

SOURCE=.\opcquality.h
# End Source File
# Begin Source File

SOURCE=.\opctestclient.h
# End Source File
# Begin Source File

SOURCE=.\safearray.h
# End Source File
# Begin Source File

SOURCE=.\safelock.h
# End Source File
# Begin Source File

SOURCE=.\server.h
# End Source File
# Begin Source File

SOURCE=.\serverenumgroupsdlg.h
# End Source File
# Begin Source File

SOURCE=.\servergeterrorstringdlg.h
# End Source File
# Begin Source File

SOURCE=.\servergroupbynamedlg.h
# End Source File
# Begin Source File

SOURCE=.\serverpropertysheet.h
# End Source File
# Begin Source File

SOURCE=.\shutdownsink.h
# End Source File
# Begin Source File

SOURCE=.\smarttooltip.h
# End Source File
# Begin Source File

SOURCE=.\statusbartext.h
# End Source File
# Begin Source File

SOURCE=.\stdafx.h
# End Source File
# Begin Source File

SOURCE=.\timestmp.h
# End Source File
# Begin Source File

SOURCE=.\updateintervaldlg.h
# End Source File
# Begin Source File

SOURCE=.\versioninfo.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;cnt;rtf;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\opctestclient.ico
# End Source File
# Begin Source File

SOURCE=.\opctestclient.rc
# End Source File
# Begin Source File

SOURCE=.\resource.h
# End Source File
# End Group
# Begin Group "Bitmaps"

# PROP Default_Filter "bmp"
# Begin Source File

SOURCE=.\res\checkboxes.bmp
# End Source File
# Begin Source File

SOURCE=.\res\components.bmp
# End Source File
# Begin Source File

SOURCE=.\res\deleteitem.bmp
# End Source File
# Begin Source File

SOURCE=.\res\deleteitemgray.bmp
# End Source File
# Begin Source File

SOURCE=.\res\dupitem.bmp
# End Source File
# Begin Source File

SOURCE=.\res\dupitemgray.bmp
# End Source File
# Begin Source File

SOURCE=.\res\eventimages.bmp
# End Source File
# Begin Source File

SOURCE=.\res\groupimages.bmp
# End Source File
# Begin Source File

SOURCE=.\res\itemimages.bmp
# End Source File
# Begin Source File

SOURCE=.\res\newitem.bmp
# End Source File
# Begin Source File

SOURCE=.\res\newitemgray.bmp
# End Source File
# Begin Source File

SOURCE=.\res\next.bmp
# End Source File
# Begin Source File

SOURCE=.\res\nextgray.bmp
# End Source File
# Begin Source File

SOURCE=.\res\previous.bmp
# End Source File
# Begin Source File

SOURCE=.\res\previousgray.bmp
# End Source File
# Begin Source File

SOURCE=.\res\toolbar.bmp
# End Source File
# Begin Source File

SOURCE=.\res\validateitem.bmp
# End Source File
# Begin Source File

SOURCE=.\res\validateitemgray.bmp
# End Source File
# End Group
# Begin Source File

SOURCE=.\README.txt
# End Source File
# End Target
# End Project
