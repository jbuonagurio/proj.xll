## Builds PROJ.XLL with Visual C++
## "%VCINSTALLDIR%\vcvarsall.bat" x86
## nmake -f Makefile

# /W3           display level 1, level 2 and level 3 (production quality) warnings
# /WX           treat all compiler warnings as errors
# /EHsc         exception-handling model: catch C++ exceptions only
# /sdl          enable additional security checks
# /MDd          create a debug multithreaded DLL using MSVCRTD.lib
# /Zc:wchar_t   parse wchar_t as a built-in type
# /Zc:inline    remove unreferenced COMDAT
# /Fo           create object files
# /Gd           use the __cdecl calling convention
# /TC           files are C source files

CFLAGS = /W3 /WX /EHsc /sdl /MDd /Zc:wchar_t /Zc:inline /Fo"debug\\" /Gd /TC

# Installation Directory
INSTDIR=%USERPROFILE%\AppData\Roaming\Microsoft\AddIns\

# Microsoft Excel 2013 XLL SDK Paths
XLLSDK_DIR=C:\2013 Office System Developer Resources\Excel2013XLLSDK

XLLSDK_LIB=$(XLLSDK_DIR)\LIB

XLLSDK_INCLUDE=$(XLLSDK_DIR)\INCLUDE

# PROJ.4 Paths
PROJ_LIB=C:\PROJ\lib\

PROJ_INCLUDE=$(PROJ_LIB)\..\include

# Project
XLLNAME = proj.xll

SOURCES = addin.c util.c

OBJECTS = $(SOURCES:.c=.obj)

all: xll

xll: $(OBJECTS)
	link /out:debug\$(XLLNAME) /libpath:"$(XLLSDK_LIB)" /libpath:"$(PROJ_LIB)" \
	XLCALL32.LIB frmwrk32.lib proj.lib user32.lib /nodefaultlib:LIBCMT \
	/def:"addin.def" /machine:x86 /dll debug\addin.obj debug\util.obj

.c.obj:
	cl /c $(CFLAGS) /I"$(XLLSDK_INCLUDE)" /I"$(PROJ_INCLUDE)" $<

clean:
	del debug\*.xll
	del debug\*.obj
    del debug\*.exp
    del debug\*.lib

install: xll
	copy /Y debug\$(XLLNAME) "$(INSTDIR)"

