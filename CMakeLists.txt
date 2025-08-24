cmake_minimum_required(VERSION 3.31)
project(AutoHotkey-HID LANGUAGES CXX)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

find_package(Boost REQUIRED COMPONENTS utility locale system)
find_package(hidapi CONFIG REQUIRED)

file(DOWNLOAD
  https://raw.githubusercontent.com/acdemiralp/hid.hpp/refs/heads/main/include/hid/hid.hpp
  ${CMAKE_CURRENT_BINARY_DIR}/hid/hid.hpp
  EXPECTED_HASH SHA256=bf5b0129c5f978a3dba484b470c232e2ed41b70500e2eb891d48711c29ef5373)

file(GLOB_RECURSE src_files CONFIGURE_DEPENDS
  "src/AutoHotkey-HID/*.h"
  "src/AutoHotkey-HID/*.cpp"
  "src/AutoHotkey-HID/*.rc")
add_library(AutoHotkey-HID MODULE ${src_files})
target_include_directories(AutoHotkey-HID PRIVATE src ${CMAKE_CURRENT_BINARY_DIR}/hid)
target_compile_features(AutoHotkey-HID PUBLIC cxx_std_23)
target_compile_options(AutoHotkey-HID PUBLIC
  /utf-8
  /external:anglebrackets
  /external:W0
  /WX
  /W4
  /analyze
  /analyze:external-)
target_compile_definitions(AutoHotkey-HID PUBLIC
  UNICODE
  WIN32_LEAN_AND_MEAN
  VC_EXTRALEAN
  NOMINMAX
  _WINDLL
  $<$<CONFIG:Debug>:DEBUG>)
target_link_libraries(AutoHotkey-HID PRIVATE Boost::utility Boost::locale Boost::system
                                             hidapi::hidapi hidapi::winapi hidapi::include)
target_precompile_headers(AutoHotkey-HID PRIVATE src/AutoHotkey-HID/PrecompiledHeader.h)

# Compile IDL file
find_program(MIDL midl.exe REQUIRED)
set(MIDL_OUT_DIR ${CMAKE_CURRENT_BINARY_DIR}/midl/AutoHotkey-HID/Gen)
file(MAKE_DIRECTORY ${MIDL_OUT_DIR})
set(IDL_FILE ${CMAKE_CURRENT_SOURCE_DIR}/src/AutoHotkey-HID/HID.idl)
set(MIDL_SOURCES ${MIDL_OUT_DIR}/HID_i.cpp ${MIDL_OUT_DIR}/HID_i.h ${MIDL_OUT_DIR}/HID.tlb)
add_custom_command(
  OUTPUT ${MIDL_SOURCES}
  COMMAND ${MIDL}
          /nologo
          /char signed
          /env x64
          /target NT61
          /h HID_i.h
          /iid HID_i.cpp
          /out ${MIDL_OUT_DIR}
          ${IDL_FILE}
  DEPENDS ${IDL_FILE}
  WORKING_DIRECTORY ${MIDL_OUT_DIR}
  VERBATIM)
target_sources(AutoHotkey-HID PRIVATE ${MIDL_SOURCES})
target_include_directories(AutoHotkey-HID PRIVATE ${MIDL_OUT_DIR}/../..)
