#  :copyright: 2026. Grégory Lerbret.
#  :license: SPDX-License-Identifier: BSD-2-Clause

# CMake file for Xlsxwriter++
# ============================
#
# To build Xlsxwriter++ with CMake move to a custom directory, ideally outside
# the root of the source tree, and type `cmake $XWPP_SOURCE_LOCATION $FLAGS`,
# where `XWPP_SOURCE_LOCATION` is the path to the Xlsxwriter++ project, and
# `FLAGS` are custom flags to pass to the compiler.
#
# For example, in the project directory, you can build Xlsxwriter++ as follows:
#
# ```
#     mkdir build
#     cd build
#
#     cmake .. -DCMAKE_BUILD_TYPE=Release
#     cmake --build . --config Release
#
# ```
#
# To build examples:
# ```
#     mkdir build
#     cd build
#
#     cmake .. -DBUILD_EXAMPLES=ON -DCMAKE_BUILD_TYPE=Release
#     cmake --build . --config Release
# ```
#
# Or to build and run the tests (this can take several minutes):
#
# ```
#     mkdir build
#     cd build
#
#     cmake .. -DBUILD_TESTS=ON -DCMAKE_BUILD_TYPE=Release
#     cmake --build . --config Release
# ```
#
# If the tests run successfully, you can install the library as follows:
#
# ```
#     cmake --build . --config Release --target install
# ```
#
# The Xlsxwriter++ CMake options are shown below.

# ---------------------
# Project configuration
# ---------------------
cmake_minimum_required(VERSION 3.16)

if(POLICY CMP0167)
  cmake_policy(SET CMP0167 NEW)
endif()

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_EXTENSIONS OFF)
set(CMAKE_CXX_STANDARD_REQUIRED True)
set(CMAKE_WINDOWS_EXPORT_ALL_SYMBOLS ON)

set(XLSXPP_PROJECT_NAME
  "xlsxwriterpp"
  CACHE STRING
  "Xlsxwriter++ is a C++ library for creating new Excel XLSX files"
)

project(${XLSXPP_PROJECT_NAME} C CXX)

if(MSVC)
  # warning level 4
  add_compile_options(/W4 /WX)
else()
  # additional warnings
  add_compile_options(-Wall -Wextra -Wpedantic -Werror -Wno-missing-field-initializers)
endif()

# -------------------
# Compilation options
# -------------------

# `USE_SYSTEM_MINIZIP`
#
# Xlsxwriter++ uses the `minizip` component of `zlib` to create the xlsx zip
# file container. The vendored source files for `minizip` are included in the
# src tree of Xlsxwriter++ and are statically linked by default.
#
# You can use this option if you have the `minizip` library already installed on
# your system and prefer to dynamically link against it.
#
# To enable this option pass `-DUSE_SYSTEM_MINIZIP=ON` during configuration.
option(
  USE_SYSTEM_MINIZIP
  "Use system minizip library instead of the vendored copy"
  OFF
)

# `BUILD_TESTS`
#
# Compile the unit and function tests for Xlsxwriter++. This functional tests
# require the Boost.Test library. The tests can be run after compilation.
#
# To enable this option pass `-DBUILD_TESTS=ON` during configuration.
option(
  BUILD_TESTS
  "Build the Xlsxwriter++ unit and functional tests (requires Boost.Test)"
  OFF
)

# `BUILD_EXAMPLES`
#
# Compile the Xlsxwriter++ example programs.
#
# To enable this option pass `-DBUILD_EXAMPLES=ON` during configuration.
option(BUILD_EXAMPLES "Build Xlsxwriter++ examples" OFF)

# `USE_STATIC_MSVC_RUNTIME`
#
# Compile as a static library with the static MSVC runtime library.
#
# To enable this option pass `-DUSE_STATIC_MSVC_RUNTIME=ON` during configuration.
if(MSVC)
  option(USE_STATIC_MSVC_RUNTIME "Use the static runtime library" OFF)
endif()

# `ZLIB_ROOT`
#
# The `ZLIB `root directory can be specified either through an environment
# variable (`export ZLIB_ROOT=/usr/include`) or using a flag with CMake
# (`-DZLIB_ROOT:STRING=/usr/include`). This sets the preferred search path for
# the ZLIB installation.
set(ZLIB_ROOT "" CACHE STRING "Optional root for the ZLIB installation")
if(DEFINED ENV{ZLIB_ROOT})
  set(ZLIB_ROOT $ENV{ZLIB_ROOT})
endif()

# -------------------------
# Configure the compilation
# -------------------------

if(USE_SYSTEM_MINIZIP)
  list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS USE_SYSTEM_MINIZIP)
endif()

if(NOT BUILD_SHARED_LIBS)
  if(UNIX)
    set(CMAKE_POSITION_INDEPENDENT_CODE ON)
  elseif(MINGW OR MSYS)
    set(CMAKE_C_FLAGS
      "${CMAKE_C_FLAGS} -static -static-libgcc -Wno-char-subscripts -Wno-long-long"
    )
    list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS USE_FILE32API)
  elseif(MSVC)
    set(CMAKE_C_FLAGS_DEBUG
      "${CMAKE_C_FLAGS_DEBUG} /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
    )
    set(CMAKE_C_FLAGS_RELEASE
      "${CMAKE_C_FLAGS_RELEASE} /Ox /Zi /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
    )
    set(CMAKE_C_FLAGS_MINSIZEREL
      "${CMAKE_C_FLAGS_MINSIZEREL} /Zi /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
    )
    set(CMAKE_C_FLAGS_RELWITHDEBINFO
      "${CMAKE_C_FLAGS_RELWITHDEBINFO} /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
    )
  endif()
endif()

if(MSVC AND USE_STATIC_MSVC_RUNTIME)
  foreach(
    flag_var
    CMAKE_C_FLAGS
    CMAKE_C_FLAGS_DEBUG
    CMAKE_C_FLAGS_RELEASE
    CMAKE_C_FLAGS_MINSIZEREL
    CMAKE_C_FLAGS_RELWITHDEBINFO
  )
    if(${flag_var} MATCHES "/MD")
      string(REPLACE "/MD" "/MT" ${flag_var} "${${flag_var}}")
    endif()
  endforeach()
endif()

file(READ "include/xlsxwriterpp.h" ver)
string(REGEX MATCH "XWPP_VERSION \"([^\"]+)\"" _ ${ver})
set(VERSION ${CMAKE_MATCH_1})
string(REGEX MATCH "XWPP_SOVERSION \"([^\"]+)\"" _ ${ver})
set(SOVERSION ${CMAKE_MATCH_1})

set(ENABLED_OPTIONS "zlib")
if(USE_SYSTEM_MINIZIP)
  string(APPEND ENABLED_OPTIONS " minizip")
endif()
string(APPEND ENABLED_OPTIONS " libcrypto")

# ----------------
# Set the includes
# ----------------
enable_language(CXX)
list(APPEND CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)

# Set the zlib includes.
find_package(ZLIB "1.2.8" REQUIRED)
list(APPEND XWPP_PRIVATE_INCLUDE_DIRS ${ZLIB_INCLUDE_DIRS})

# Set the minizip includes.
if(USE_SYSTEM_MINIZIP)
  if(MSVC)
    find_package(MINIZIP NAMES unofficial-minizip REQUIRED)
    set(MINIZIP_LIBRARIES unofficial::minizip::minizip)
  else()
    find_package(PkgConfig REQUIRED)
    pkg_check_modules(MINIZIP minizip)
    list(APPEND XWPP_PRIVATE_INCLUDE_DIRS ${MINIZIP_INCLUDE_DIRS}/..)
  endif()
endif()

# ----------------------------
# Set the library dependencies
# ----------------------------
list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS NOCRYPT NOUNCRYPT)

# Ensure CRT Secure warnings are disabled.
if(MSVC)
  list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS _CRT_SECURE_NO_WARNINGS)
endif()

# Set the source files.
file(GLOB XWPP_SOURCES src/*.cpp)
file(GLOB_RECURSE XWPP_HEADERS RELATIVE include *.h)

# If not using the system minizip, add the vendored minizip source files.
if(NOT USE_SYSTEM_MINIZIP)
  list(
    APPEND XWPP_SOURCES
    third_party/minizip/ioapi.c
    third_party/minizip/zip.c
  )
  if(MSVC)
    add_compile_options(/wd4100 /wd4189)
    list(APPEND XWPP_SOURCES third_party/minizip/iowin32.c)
  endif()
endif()

# Set project metadata.
set(XWPP_PROJECT_DIR "${CMAKE_CURRENT_SOURCE_DIR}")
set(XWPP_LIB_DIR "${XWPP_PROJECT_DIR}/lib")
add_library(${PROJECT_NAME} "")
set_target_properties(${PROJECT_NAME} PROPERTIES SOVERSION ${SOVERSION})
target_sources(${PROJECT_NAME} PRIVATE ${XWPP_SOURCES} PUBLIC ${XWPP_HEADERS})

target_link_libraries(${PROJECT_NAME} PRIVATE ZLIB::ZLIB)

if(MINIZIP_LINK_LIBRARIES)
  target_link_libraries(${PROJECT_NAME} PRIVATE ${MINIZIP_LINK_LIBRARIES})
else()
  target_link_libraries(${PROJECT_NAME} PRIVATE ${MINIZIP_LIBRARIES})
endif()

target_link_libraries(${PROJECT_NAME} PRIVATE ${LIB_CRYPTO} ${OPENSSL_CRYPTO_LIBRARY})

target_compile_definitions(
    ${PROJECT_NAME}
    PRIVATE ${XWPP_PRIVATE_COMPILE_DEFINITIONS}
)

# Ensure MSVC supports /utf-8.
if(MSVC AND MSVC_VERSION GREATER_EQUAL 1900)
  target_compile_options(${PROJECT_NAME} PRIVATE /utf-8)
endif()

if(WINDOWSSTORE)
  target_compile_definitions(
    ${PROJECT_NAME}
    PRIVATE -DIOWIN32_USING_WINRT_API
  )
endif()

target_include_directories(
  ${PROJECT_NAME}
  PRIVATE ${XWPP_PRIVATE_INCLUDE_DIRS}
  PUBLIC include
)

# ----------------
# Set up the tests
# ------ ---------
if(BUILD_TESTS)
  find_package(Boost 1.83 REQUIRED COMPONENTS unit_test_framework)

  if(MSVC)
    # Disable warnings C4005 caused by BOOST_TEST_DYN_LINK
    add_compile_options(/wd4005)
  endif()

  MACRO(SUBDIRLIST result curdir)
    FILE(GLOB children ${curdir}/*)
    SET(dirlist "")
    FOREACH(child ${children})
      IF(IS_DIRECTORY ${child})
        get_filename_component(basename ${child} NAME)
        LIST(APPEND dirlist ${basename})
      ENDIF()
    ENDFOREACH()
    SET(${result} ${dirlist})
  ENDMACRO()

  # Unit test
  SUBDIRLIST(SUBDIRS "test/unit")
  FOREACH(subdir ${SUBDIRS})
    set(test_exec_name "test_${subdir}")
    file(GLOB XWPP_SOURCES test/unit/${subdir}/*.cpp)
    add_executable(${test_exec_name} ${XWPP_SOURCES})
    target_link_libraries(${test_exec_name} ${PROJECT_NAME} ${Boost_LIBRARIES})
    set_target_properties(${test_exec_name} PROPERTIES RUNTIME_OUTPUT_DIRECTORY $<1:test/unit>)
  ENDFOREACH()

  file(GLOB XWPP_FUNCTIONAL_SOURCES test/functional/src/*.cpp)
  foreach(source ${XWPP_FUNCTIONAL_SOURCES})
    get_filename_component(basename ${source} NAME_WE)
    add_executable(${basename} ${source})
    target_link_libraries(${basename} ${PROJECT_NAME})
    set_target_properties(${basename} PROPERTIES RUNTIME_OUTPUT_DIRECTORY $<1:test/functional/src>)
  endforeach(source)

  # Custom target and commands to copy necessary files for testing
  # They use fake files to force copy
  add_custom_target(
    custom_test_target_runs_always ALL
    DEPENDS ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_xlsx ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_py ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_images
  )

  set_source_files_properties(${CMAKE_CURRENT_BINARY_DIR}/__test_copy_xlsx PROPERTIES SYMBOLIC 1)
  add_custom_command(
    OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_xlsx
    COMMAND ${CMAKE_COMMAND} -E copy_directory_if_different
            ${CMAKE_SOURCE_DIR}/test/functional/xlsx_files
            ${CMAKE_CURRENT_BINARY_DIR}/test/functional/xlsx_files)

  set_source_files_properties(${CMAKE_CURRENT_BINARY_DIR}/__test_copy_images PROPERTIES SYMBOLIC 1)
  add_custom_command(
    OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_images
    COMMAND ${CMAKE_COMMAND} -E copy_directory_if_different
            ${CMAKE_SOURCE_DIR}/test/functional/src/images
            ${CMAKE_CURRENT_BINARY_DIR}/test/functional/src/images)

  file(GLOB PYTHON_TEST_SRC ${CMAKE_SOURCE_DIR}/test/functional/*.py)
  set_source_files_properties(${CMAKE_CURRENT_BINARY_DIR}/__test_copy_py PROPERTIES SYMBOLIC 1)
  add_custom_command(
    OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/__test_copy_py
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            ${PYTHON_TEST_SRC}
            ${CMAKE_CURRENT_BINARY_DIR}/test/functional)
endif()

# ------------------------------------
# Compile and run the example programs
# ------------------------------------

file(GLOB XWPP_EXAMPLE_SOURCES examples/*.cpp)

if(BUILD_EXAMPLES)
  foreach(source ${XWPP_EXAMPLE_SOURCES})
    get_filename_component(basename ${source} NAME_WE)
    add_executable(${basename} ${source})
    target_link_libraries(${basename} ${PROJECT_NAME})
    set_target_properties(${basename} PROPERTIES RUNTIME_OUTPUT_DIRECTORY $<1:examples>)
  endforeach(source)

  # Custom target and commands to copy necessary files for examples
  # They use fake files to force copy
  add_custom_target(
    custom_example_target_runs_always ALL
    DEPENDS ${CMAKE_CURRENT_BINARY_DIR}/__example_copy_images  ${CMAKE_CURRENT_BINARY_DIR}/__example_copy_bin
  )

  file(MAKE_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/examples RESULT res)

  file(GLOB EXAMPLES_IMAGES ${CMAKE_SOURCE_DIR}/examples/*.png)
  set_source_files_properties(${CMAKE_CURRENT_BINARY_DIR}/__example_copy_images PROPERTIES SYMBOLIC 1)
  add_custom_command(
    OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/__example_copy_images
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            ${EXAMPLES_IMAGES}
            ${CMAKE_CURRENT_BINARY_DIR}/examples)

  file(GLOB EXAMPLES_BIN ${CMAKE_SOURCE_DIR}/examples/*.bin)
  set_source_files_properties(${CMAKE_CURRENT_BINARY_DIR}/__example_copy_bin PROPERTIES SYMBOLIC 1)
  add_custom_command(
    OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/__example_copy_bin
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            ${EXAMPLES_BIN}
            ${CMAKE_CURRENT_BINARY_DIR}/examples)
endif()

# -------------------
# Install the library
# -------------------
include(GNUInstallDirs)

install(
  TARGETS ${PROJECT_NAME}
  LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
  ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
  RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)

install(FILES include/xlsxwriterpp.h DESTINATION include)

install(
  DIRECTORY include/xlsxwriter
  DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
  FILES_MATCHING
  PATTERN "*.h"
)
