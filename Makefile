## Builds PROJ.XLL with Visual C++
## "%VCINSTALLDIR%\vcvarsall.bat" x86
## nmake -f Makefile

# /nologo       suppress startup banner
# /O2           minimize size, maximize speed
# /W3           display level 1, level 2 and level 3 (production quality) warnings
# /WX           treat all compiler warnings as errors
# /EHsc         exception-handling model: catch C++ exceptions only
# /MT           use the multithreaded static CRT (LIBCMT)
# /Zc:wchar_t   parse wchar_t as a built-in type
# /Zc:inline    remove unreferenced COMDAT
# /Gd           use the __cdecl calling convention
# /TC           files are C source files

CFLAGS = /nologo /O2 /W3 /WX /EHsc /MT /Zc:wchar_t /Zc:inline /Gd /TC

# Build Artifacts
BUILDDIR = ..\build\

# Installation Directory
INSTDIR = %USERPROFILE%\AppData\Roaming\Microsoft\AddIns\

# Microsoft Excel 2013 XLL SDK Paths
XLLSDK_DIR = C:\2013 Office System Developer Resources\Excel2013XLLSDK
XLLSDK_LIBPATH = $(XLLSDK_DIR)\LIB
XLLSDK_INCLUDE = $(XLLSDK_DIR)\INCLUDE
XLLSDK_FRAMEWRK_LIBPATH = $(XLLSDK_DIR)\SAMPLES\FRAMEWRK\RELEASE

# PROJ Paths
PROJ_DIR = C:\PROJ
PROJ_LIBPATH = $(PROJ_DIR)\lib
PROJ_INCLUDE = $(PROJ_DIR)\include

# Target
XLLNAME = proj.xll

SOURCES = addin.c util.c

OBJECTS = $(SOURCES:.c=.obj)

all: xll

"$(BUILDDIR)":
	if not exist "$(BUILDDIR)" mkdir "$(BUILDDIR)"

xll: "$(BUILDDIR)" $(OBJECTS)
	link /nologo /out:"$(BUILDDIR)\$(XLLNAME)" /libpath:"$(XLLSDK_LIBPATH)" /libpath:"$(XLLSDK_FRAMEWRK_LIBPATH)" /libpath:"$(PROJ_LIBPATH)" \
	XLCALL32.LIB frmwrk32.lib proj.lib user32.lib \
	/def:"addin.def" /machine:x86 /dll "$(BUILDDIR)\addin.obj" "$(BUILDDIR)\util.obj"

.c.obj:
	cl /c $(CFLAGS) /I"$(XLLSDK_INCLUDE)" /I"$(PROJ_INCLUDE)" /Fo"$(BUILDDIR)\\" $<

clean:
	del "$(BUILDDIR)\*.xll"
	del "$(BUILDDIR)\*.obj"
    del "$(BUILDDIR)\*.exp"
    del "$(BUILDDIR)\*.lib"

install: xll
	copy /Y $(BUILDDIR)\$(XLLNAME) "$(INSTDIR)"

