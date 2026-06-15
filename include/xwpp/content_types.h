/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * content_types - A Xlsxwriter++ library for creating Excel XLSX
 *                 content_types files.
 *
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
  void add_shared_strings();
  void add_custom_properties();
  void add_vml_name();
  void add_comment_name(std::string_view name);
  void add_drawing_name(std::string_view name);
  void add_metadata();
  void add_rich_value();
  void add_chart_name(std::string_view name);

  [[nodiscard]] std::string assemble_xml_file() const;

  static const std::string APP_PACKAGE;
  static const std::string APP_DOCUMENT;
  static const std::string APP_MSEXCEL;

private:
  std::vector<std::tuple<std::string, std::string>> default_types_;
  std::vector<std::tuple<std::string, std::string>> overrides_;
};

/// lxw_content_types *lxw_content_types_new(void);
/// void lxw_content_types_free(lxw_content_types *content_types);
/// void lxw_ct_add_chartsheet_name(lxw_content_types *content_types,
///                                 const char *name);
/// void lxw_ct_add_table_name(lxw_content_types *content_types,
///                            const char *name);
/// void lxw_ct_add_calc_chain(lxw_content_types *content_types);
/// STATIC void _content_types_xml_declaration(lxw_content_types *self);
/// STATIC void _write_default(lxw_content_types *self, const char *ext,
///                            const char *type);
/// STATIC void _write_override(lxw_content_types *self, const char *part_name,
///                             const char *type);

}

#endif
