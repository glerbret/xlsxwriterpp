/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 * styles - A Xlsxwriter++ library for creating Excel XLSX styles files.
 *
 */

#ifndef XWPP_STYLES_H
#define XWPP_STYLES_H

#include "xwpp/format.h"

#include <vector>

namespace xwpp
{

class styles_t
{
public:
  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_style_sheet() const;
  [[nodiscard]] std::string write_num_fmts() const;
  [[nodiscard]] std::string write_fonts() const;
  [[nodiscard]] std::string write_default_fill(const std::string& pattern) const;
  [[nodiscard]] std::string write_fills() const;
  [[nodiscard]] std::string write_borders() const;
  [[nodiscard]] std::string write_style_xf(bool has_hyperlink, uint16_t font_id) const;
  [[nodiscard]] std::string write_cell_style_xfs() const;
  [[nodiscard]] std::string write_cell_xfs() const;
  [[nodiscard]] std::string write_cell_style(const std::string& name, uint8_t xf_id, uint8_t builtin_id) const;
  [[nodiscard]] std::string write_cell_styles() const;
  [[nodiscard]] std::string write_dxfs() const;
  [[nodiscard]] std::string write_table_styles() const;

  ///  uint32_t font_count;
  ///  uint32_t xf_count;
  ///  uint32_t dxf_count;
  ///  uint32_t num_format_count;
  ///  uint32_t border_count;
  ///  uint32_t fill_count;
  std::vector<format_t> xf_formats_;
  std::vector<format_t> dxf_formats_;
  ///  uint8_t has_hyperlink;
  ///  uint16_t hyperlink_font_id;
  ///  uint8_t has_comments;
};

/// void lxw_styles_write_string_fragment(lxw_styles *self, const char *string);
/// void lxw_styles_write_rich_font(lxw_styles *styles, lxw_format *format);

/// STATIC void _write_font_size(lxw_styles *self, double font_size);
/// STATIC void _write_font_color_theme(lxw_styles *self, uint8_t theme);
/// STATIC void _write_font_name(lxw_styles *self, const char *font_name,
///                              uint8_t is_rich_string);
/// STATIC void _write_font_family(lxw_styles *self, uint8_t font_family);
/// STATIC void _write_font_scheme(lxw_styles *self, const char *font_scheme);
/// STATIC void _write_font(lxw_styles *self, lxw_format *format, uint8_t is_dxf,
///                         uint8_t is_rich_string);

/// STATIC void _write_border(lxw_styles *self, lxw_format *format,
///                           uint8_t is_dxf);
/// STATIC void _write_xf(lxw_styles *self, lxw_format *format);
}

#endif
