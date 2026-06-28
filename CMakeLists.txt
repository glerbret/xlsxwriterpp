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
#### If the tests run successfully, you can install the library as follows:
#
#### ```
####     cmake --build . --config Release --target install
#### ```
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
if (MSVC)
  # warning level 4
  add_compile_options(/W4)
else()
  # additional warnings
  add_compile_options(-Wall -Wextra -Wpedantic -Wno-comment -Wno-missing-field-initializers)
endif()

set(XLSXPP_PROJECT_NAME
  "xlsxwriterpp"
  CACHE STRING
  "Xlsxwriter++ is a C++ library for creating new Excel XLSX files"
)

project(${XLSXPP_PROJECT_NAME} C CXX)
### enable_testing()

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
# option(
#     BUILD_TESTS
#     "Build the Xlsxwriter++ unit and functional tests (requires Boost.Test)"
#     OFF
# )

# TODO XWPP_TARGET_BIG_ENDIAN

# `BUILD_EXAMPLES`
#
# Compile the Xlsxwriter++ example programs.
#
# To enable this option pass `-DBUILD_EXAMPLES=ON` during configuration.
option(BUILD_EXAMPLES "Build Xlsxwriter++ examples" OFF)

# `BUILD_FUZZERS`
#
# Compile the fuzzer harnesses.
#
# To enable this option pass `-DBUILD_FUZZERS=ON` during configuration.
### option(BUILD_FUZZERS "Build harness(es) for fuzzing" OFF)

# `USE_STATIC_MSVC_RUNTIME`
#
# Compile as a static library with the static MSVC runtime library.
#
# To enable this option pass `-DUSE_STATIC_MSVC_RUNTIME=ON` during configuration.
# if(MSVC)
#     option(USE_STATIC_MSVC_RUNTIME "Use the static runtime library" OFF)
# endif()

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

## TODO Replace MD5_Init and co and remove this
# For openSSL use
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-deprecated-declarations")

### if(NOT BUILD_SHARED_LIBS)
###     if(UNIX)
###         set(CMAKE_POSITION_INDEPENDENT_CODE ON)
###     elseif(MINGW OR MSYS)
###         set(CMAKE_C_FLAGS
###             "${CMAKE_C_FLAGS} -static -static-libgcc -Wno-char-subscripts -Wno-long-long"
###         )
###         list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS USE_FILE32API)
###     elseif(MSVC)
###         set(CMAKE_C_FLAGS_DEBUG
###             "${CMAKE_C_FLAGS_DEBUG} /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
###         )
###         set(CMAKE_C_FLAGS_RELEASE
###             "${CMAKE_C_FLAGS_RELEASE} /Ox /Zi /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
###         )
###         set(CMAKE_C_FLAGS_MINSIZEREL
###             "${CMAKE_C_FLAGS_MINSIZEREL} /Zi /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
###         )
###         set(CMAKE_C_FLAGS_RELWITHDEBINFO
###             "${CMAKE_C_FLAGS_RELWITHDEBINFO} /Fd\"${CMAKE_BINARY_DIR}/${PROJECT_NAME}.pdb\""
###         )
###     endif()
### endif()

### if(MSVC AND USE_STATIC_MSVC_RUNTIME)
###     foreach(
###         flag_var
###         CMAKE_C_FLAGS
###         CMAKE_C_FLAGS_DEBUG
###         CMAKE_C_FLAGS_RELEASE
###         CMAKE_C_FLAGS_MINSIZEREL
###         CMAKE_C_FLAGS_RELWITHDEBINFO
###     )
###         if(${flag_var} MATCHES "/MD")
###             string(REPLACE "/MD" "/MT" ${flag_var} "${${flag_var}}")
###         endif()
###     endforeach()
### endif()

# -----------------------------
# Create the xlsxwriterpp.pc file
# -----------------------------
set(PREFIX ${CMAKE_INSTALL_PREFIX})

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

# Expand out the xlsxwriterpp.pc file.
### configure_file(dev/release/pkg-config.txt xlsxwriterpp.pc @ONLY)

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

# Set the openssl includes.
find_package(OpenSSL REQUIRED)
include_directories(${OPENSSL_INCLUDE_DIR})

# ----------------------------
# Set the library dependencies
# ----------------------------
### list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS NOCRYPT NOUNCRYPT)

# Ensure CRT Secure warnings are disabled.
### if(MSVC)
###     list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS _CRT_SECURE_NO_WARNINGS)
### endif()

# Ensure "TESTING" macro is defined if building tests.
### if(BUILD_TESTS)
###     list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS TESTING)
### endif()

# Define "XWPP_BIG_ENDIAN" macro on big-endian architectures.
### include(TestBigEndian)
### test_big_endian(XWPP_TARGET_BIG_ENDIAN)
### if(XWPP_TARGET_BIG_ENDIAN)
###     list(APPEND XWPP_PRIVATE_COMPILE_DEFINITIONS XWPP_BIG_ENDIAN)
### endif()

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

### target_compile_definitions(
###     ${PROJECT_NAME}
###     PRIVATE ${XWPP_PRIVATE_COMPILE_DEFINITIONS}
### )

# Ensure MSVC supports /utf-8.
### if(MSVC AND MSVC_VERSION GREATER_EQUAL 1900)
###     target_compile_options(${PROJECT_NAME} PRIVATE /utf-8)
### endif()

### if(WINDOWSSTORE)
###     target_compile_definitions(
###         ${PROJECT_NAME}
###         PRIVATE -DIOWIN32_USING_WINRT_API
###     )
### endif()

target_include_directories(
  ${PROJECT_NAME}
  PRIVATE ${XWPP_PRIVATE_INCLUDE_DIRS}
  PUBLIC include
)

# ----------------
# Set up the tests
# ------ ---------
if(NOT MSVC)



# Skip unit tests on Windows.
  set(XWPP_UNIT_SOURCES
###         test/unit/test_all.c
###         ${XWPP_UTILITY_SOURCES}
###         ${XWPP_XMLWRITER_SOURCES}
###         ${XWPP_WORKSHEET_SOURCES}
###         ${XWPP_SST_SOURCES}
###         ${XWPP_WORKBOOK_SOURCES}
###         ${XWPP_APP_SOURCES}
###         ${XWPP_CONTENTTYPES_SOURCES}
###         ${XWPP_CORE_SOURCES}
###         ${XWPP_RELATIONSHIPS_SOURCES}
###         ${XWPP_FORMAT_SOURCES}
###         ${XWPP_STYLES_SOURCES}
###         ${XWPP_DRAWING_SOURCES}
###         ${XWPP_CHART_SOURCES}
###         ${XWPP_CUSTOM_SOURCES}
    )
### else()
###     set(XWPP_UNIT_SOURCES test/cpp/test_compilation.cpp)
endif()

## TODO Autodetect unit tests
if(BUILD_TESTS)
  find_package(Boost 1.83 REQUIRED COMPONENTS unit_test_framework)

  file(GLOB XWPP_APP_SOURCES test/unit/app/*.cpp)
  add_executable(test_app ${XWPP_APP_SOURCES})
  target_link_libraries(test_app ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties( test_app PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_CORE_SOURCES test/unit/core/*.cpp)
  add_executable(test_core ${XWPP_CORE_SOURCES})
  target_link_libraries(test_core ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties( test_core PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_CONTENTTYPES_SOURCES test/unit/content_types/*.cpp)
  add_executable(test_content_types ${XWPP_CONTENTTYPES_SOURCES})
  target_link_libraries(test_content_types ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties( test_content_types PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_RELATIONSHIPS_SOURCES test/unit/relationships/*.cpp)
  add_executable(test_relationships ${XWPP_RELATIONSHIPS_SOURCES})
  target_link_libraries(test_relationships ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties(test_relationships PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_SST_SOURCES test/unit/sst/*.cpp)
  add_executable(test_sst ${XWPP_SST_SOURCES})
  target_link_libraries(test_sst ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties(test_sst PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_UTILITY_SOURCES test/unit/utility/*.cpp)
  add_executable(test_utility ${XWPP_UTILITY_SOURCES})
  target_link_libraries(test_utility ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties(test_utility PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_WORKBOOK_SOURCES test/unit/workbook/*.cpp)
  add_executable(test_workbook ${XWPP_WORKBOOK_SOURCES})
  target_link_libraries(test_workbook ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties(test_workbook PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

  file(GLOB XWPP_XMLWRITER_SOURCES test/unit/xmlwriter/*.cpp)
  add_executable(test_xmlwriter ${XWPP_XMLWRITER_SOURCES})
  target_link_libraries(test_xmlwriter ${PROJECT_NAME} ${Boost_LIBRARIES})
  set_target_properties(test_xmlwriter PROPERTIES RUNTIME_OUTPUT_DIRECTORY "tests")

    # Unit tests.
###     create_test(XWPP_UNIT_SOURCES unit)

    # Functional tests.
###     find_package(Python COMPONENTS Interpreter REQUIRED)
###     find_program(Pytest_EXECUTABLE NAMES pytest)

###     if(NOT Pytest_EXECUTABLE)
###         message(
###             "Please install the Python pytest library to run the functional tests:"
###         )
###         message("    pip install pytest\n")
###     endif()

###     foreach(source ${XWPP_FUNCTIONAL_SOURCES})
###         get_filename_component(basename ${source} NAME_WE)
###         add_executable(${basename} ${source})
###         target_link_libraries(${basename} xlsxwriter)
###         set_target_properties(
###             ${basename}
###             PROPERTIES RUNTIME_OUTPUT_DIRECTORY "test/functional/src"
###         )
###     endforeach(source)

###     add_custom_command(
###         TARGET xlsxwriter_unit
###         POST_BUILD
###         COMMAND
###             ${CMAKE_COMMAND} -E copy_directory
###             ${CMAKE_SOURCE_DIR}/test/functional test/functional
###     )

###     if(USE_NO_MD5)
###         add_test(
###             NAME functional
###             COMMAND pytest -v test/functional -m "not skipif"
###             WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
###         )
###     else()
###         add_test(
###             NAME functional
###             COMMAND pytest -v test/functional
###             WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
###         )
endif()

# ------------------------------------
# Compile and run the example programs
# ------------------------------------

file(GLOB XWPP_EXAMPLE_SOURCES examples/*.cpp)


# TODO Copy logo in RUNTIME_OUTPUT_DIRECTORY
if(BUILD_EXAMPLES)
  foreach(source ${XWPP_EXAMPLE_SOURCES})
    get_filename_component(basename ${source} NAME_WE)
    add_executable(${basename} ${source})
    target_link_libraries(${basename} ${PROJECT_NAME})
    set_target_properties(${basename} PROPERTIES RUNTIME_OUTPUT_DIRECTORY "examples")
  endforeach(source)
endif()

# -----------------
# Fuzzing harnesses
# -----------------
### if(BUILD_FUZZERS AND DEFINED ENV{LIB_FUZZING_ENGINE})
###     add_subdirectory(dev/fuzzing)
### endif()

# -------------------
# Install the library
# -------------------
### include(GNUInstallDirs)

### install(
###     TARGETS ${PROJECT_NAME}
###     LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
###     ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
###     RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
### )

install(FILES include/xlsxwriterpp.h DESTINATION include)

### install(
###     DIRECTORY include/xlsxwriter
###     DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
###     FILES_MATCHING
###     PATTERN "*.h"
### )
### install(
###     FILES ${CMAKE_CURRENT_BINARY_DIR}/xlsxwriterpp.pc
###     DESTINATION ${CMAKE_INSTALL_LIBDIR}/pkgconfig
### )
