/*
 * content_types - A Xlsxwriter++ library for creating Excel XLSX
 *                 content_types files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_CONTENT_TYPES_H
#define XWPP_CONTENT_TYPES_H

#include <string>
#include <string_view>
#include <tuple>
#include <vector>

namespace xwpp
{

class content_types_t
{
public:
  content_types_t();

  void add_default(std::string_view key, std::string_view value);
  void add_override(std::string_view key, std::string_view value);
  void add_worksheet_name(std::string_view name);
  void add_chartsheet_name(std::string_view name);

  void add_shared_strings();
  void add_custom_properties();
  void add_vml_name();
  void add_comment_name(std::string_view name);
  void add_drawing_name(std::string_view name);
  void add_metadata();
  void add_rich_value();
  void add_chart_name(std::string_view name);
  void add_table_name(std::string_view name);
  void add_calc_chain();

  [[nodiscard]] std::string assemble_xml_file() const;

  static const std::string APP_PACKAGE;
  static const std::string APP_DOCUMENT;
  static const std::string APP_MSEXCEL;

private:
  std::vector<std::tuple<std::string, std::string>> default_types_;
  std::vector<std::tuple<std::string, std::string>> overrides_;
};

}

#endif
