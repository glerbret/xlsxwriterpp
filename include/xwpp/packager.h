/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * packager - A Xlsxwriter++ library for creating Excel XLSX packager files.
 *
 */

#ifndef XWPP_PACKAGER_H
#define XWPP_PACKAGER_H

#include <string_view>

#ifdef USE_SYSTEM_MINIZIP
#ifdef __GNUC__
#pragma GCC system_header
#endif
#include "minizip/zip.h"
#else
#include "third_party/zip.h"
#endif

#include "workbook.h"

namespace xwpp
{

/// #define LXW_ZIP_BUFFER_SIZE (16384)

/* If zip returns a ZIP_XXX error then errno is set and we can trap that in
 * workbook.c. Otherwise return a default Xlsxwriter++ error. */
/// #define RETURN_ON_ZIP_ERROR(err, default_err)       \
///     do {                                            \
///         if (err == ZIP_ERRNO)                       \
///             return LXW_ERROR_ZIP_FILE_OPERATION;    \
///         else if (err == ZIP_PARAMERROR)             \
///             return LXW_ERROR_ZIP_PARAMETER_ERROR;   \
///         else if (err == ZIP_BADZIPFILE)             \
///             return LXW_ERROR_ZIP_BAD_ZIP_FILE;      \
///         else if (err == ZIP_INTERNALERROR)          \
///             return LXW_ERROR_ZIP_INTERNAL_ERROR;    \
///         else                                        \
///             return default_err;                     \
///     } while (0)

class packager_t
{
public:
  explicit packager_t(std::string_view filename //,
                                                /// const char *tmpdir,
                                                /// uint8_t use_zip64
  );

  // Write the xml files that make up the XLSX OPC package.
  void create_package(workbook_t& workbook);

private:
  void write_content_types_file(const workbook_t& workbook);
  void write_root_rels_file(const workbook_t& workbook);
  void write_workbook_rels_file(const workbook_t& workbook);
  void write_worksheet_files(workbook_t& workbook);
  void write_workbook_file(workbook_t& workbook);
  void write_styles_file(const workbook_t& workbook);
  void write_shared_strings_file(const workbook_t& workbook);
  void write_theme_file();
  void write_app_file(const workbook_t& workbook);
  void write_core_file(const workbook_t& workbook);
  void write_custom_file(const workbook_t& workbook);
  void write_worksheet_rels_file(const workbook_t& workbook);
  void write_vml_files(const workbook_t& workbook);
  void write_comment_files(const workbook_t& workbook);
  void write_drawing_files(const workbook_t& workbook);
  void write_drawing_rels_file(const workbook_t& workbook);
  void write_image_files(const workbook_t& workbook);
  void write_vml_drawing_rels_file(const worksheet_t& worksheet, uint32_t index);

  void add_buffer_to_zip(std::string_view buffer, const std::string& filename);
  void add_buffer_to_zip(std::vector<unsigned char> buffer, const std::string& filename);
  [[nodiscard]] uint32_t get_drawing_count(const workbook_t& workbook) const;

  ///     size_t buffer_size;
  ///     size_t output_buffer_size;
  zipFile zipfile_ = nullptr;
  zip_fileinfo zip_fileinfo_;
  std::string filename_;
  ///     const char *buffer;
  ///     char *output_buffer;
  ///     const char *tmpdir;
  ///     uint8_t use_zip64;
};

}

#endif
