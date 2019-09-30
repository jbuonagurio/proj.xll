cmake_minimum_required(VERSION 3.15)

project(projxll LANGUAGES C)

set(BUILD_SHARED_LIBS ON)

if(DEFINED ENV{VCPKG_DEFAULT_TRIPLET} AND NOT DEFINED VCPKG_TARGET_TRIPLET)
  set(VCPKG_TARGET_TRIPLET "$ENV{VCPKG_DEFAULT_TRIPLET}" CACHE STRING "")
else()
  # Use default x86 triplet with VCPKG_CRT_LINKAGE="static" if not specified.
  set(VCPKG_TARGET_TRIPLET "x86-windows-static" CACHE STRING "")
endif()

if(DEFINED ENV{VCPKG_ROOT} AND NOT DEFINED CMAKE_TOOLCHAIN_FILE)
  set(CMAKE_TOOLCHAIN_FILE "$ENV{VCPKG_ROOT}/scripts/buildsystems/vcpkg.cmake"
      CACHE STRING "")
endif()

if(CMAKE_INSTALL_PREFIX_INITIALIZED_TO_DEFAULT)
  set(CMAKE_INSTALL_PREFIX "$ENV{USERPROFILE}/AppData/Roaming/Microsoft/AddIns/"
      CACHE PATH "" FORCE)
endif()

message("CMAKE_INSTALL_PREFIX: ${CMAKE_INSTALL_PREFIX}")
message("CMAKE_TOOLCHAIN_FILE: ${CMAKE_TOOLCHAIN_FILE}")
message("VCPKG_TARGET_TRIPLET: ${VCPKG_TARGET_TRIPLET}")

#############################
# External Libraries
#############################

# Starting with PROJ version 6.1.0, proj_context_set_search_paths() has
# priority over PROJ_LIB, so require this as the minimum version.
find_package(sqlite3 CONFIG REQUIRED)
find_package(PROJ4 6.1.0 REQUIRED)

# Get the PROJ database database path.
get_filename_component(PROJ4_DB_FILE proj.db
  ABSOLUTE BASE_DIR "${PROJ4_LIBRARY_DIRS}/../share/proj/" CACHE)

find_path(XLL_INCLUDE_DIRS
  NAMES XLCALL.H
  HINTS "$ENV{SYSTEMDRIVE}/2013 Office System Developer Resources/Excel2013XLLSDK/INCLUDE")
if(NOT XLL_INCLUDE_DIRS)
  message(FATAL_ERROR "XLL include directory not found")
endif()

find_library(XLCALL32_LIBRARY
  NAMES XLCALL32
  HINTS "$ENV{SYSTEMDRIVE}/2013 Office System Developer Resources/Excel2013XLLSDK/LIB")
if(NOT XLCALL32_LIBRARY)
  message(FATAL_ERROR "XLL library not found")
endif()

find_library(FRMWRK32_LIBRARY
  NAMES frmwrk32
  HINTS "$ENV{SYSTEMDRIVE}/2013 Office System Developer Resources/Excel2013XLLSDK/LIB")
if(NOT FRMWRK32_LIBRARY)
  message(FATAL_ERROR "XLL framework library not found")
endif()

list(APPEND XLL_LIBRARIES ${XLCALL32_LIBRARY} ${FRMWRK32_LIBRARY})

#############################
# Target
#############################

# Preprocess the resource file to set the path to proj.db (PROJ_DB_FILE).
configure_file(resource.rc.in resource.rc @ONLY)

add_library(projxll SHARED
  addin.c
  addin.def
  util.c
  resource.rc
)

set_target_properties(projxll PROPERTIES
  MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>"
  OUTPUT_NAME "proj"
  SUFFIX ".xll")

target_link_libraries(projxll PRIVATE
  sqlite3
  PROJ4::proj
  ${XLL_LIBRARIES})

target_include_directories(projxll PRIVATE
  ${CMAKE_CURRENT_SOURCE_DIR}
  ${XLL_INCLUDE_DIRS})

install(TARGETS projxll)
