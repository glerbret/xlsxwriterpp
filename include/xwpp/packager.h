/*
 * packager - A Xlsxwriter++ library for creating Excel XLSX packager files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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

class packager_t
{
public:
  explicit packager_t(std::string_view filename, bool use_zip64 = false);

  // Write the xml files that make up the XLSX OPC package.
  void create_package(workbook_t& workbook);

private:
  void write_content_types_file(const workbook_t& workbook);
  void write_root_rels_file(const workbook_t& workbook);
  void write_workbook_rels_file(const workbook_t& workbook);
  void write_worksheet_files(workbook_t& workbook);
  void write_chartsheet_files(workbook_t& workbook);
  void write_workbook_file(workbook_t& workbook);
  void write_styles_file(const workbook_t& workbook);
  void write_shared_strings_file(const workbook_t& workbook);
  void write_theme_file();
  void write_app_file(const workbook_t& workbook);
  void write_core_file(const workbook_t& workbook);
  void write_custom_file(const workbook_t& workbook);
  void write_worksheet_rels_file(const workbook_t& workbook);
  void write_chartsheet_rels_file(const workbook_t& workbook);
  void write_vml_files(const workbook_t& workbook);
  void write_comment_files(const workbook_t& workbook);
  void write_drawing_files(const workbook_t& workbook);
  void write_drawing_rels_file(const workbook_t& workbook);
  void write_image_files(const workbook_t& workbook);
  void write_vml_drawing_rels_file(const worksheet_t& worksheet, uint32_t index);
  void write_rich_value_rels_file(const workbook_t& workbook);
  void write_metadata_file(const workbook_t& workbook);
  void write_rich_value_file(const workbook_t& workbook);
  void write_rich_value_rel_file(const workbook_t& workbook);
  void write_rich_value_structure_file(const workbook_t& workbook);
  void write_rich_value_types_file(const workbook_t& workbook);
  void write_chart_files(const workbook_t& workbook);
  void write_table_files(const workbook_t& workbook);
  void add_vba_project(const workbook_t& workbook);
  void add_vba_project_signature(const workbook_t& workbook);
  void write_vba_project_rels_file(const workbook_t& workbook);

  void add_buffer_to_zip(std::string_view buffer, const std::string& filename);
  void add_buffer_to_zip(std::vector<unsigned char> buffer, const std::string& filename);

  [[nodiscard]] uint32_t get_drawing_count(const workbook_t& workbook) const;
  [[nodiscard]] uint32_t get_table_count(const workbook_t& workbook) const;

  zipFile zipfile_ = nullptr;
  zip_fileinfo zip_fileinfo_;
  std::string filename_;
  bool use_zip64_ = false;
};

}

#endif
