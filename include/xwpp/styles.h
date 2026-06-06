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

#include <string>
#include <vector>

namespace xwpp
{

class style_t
{
public:
  style_t(uint32_t font_count, uint32_t border_count, const std::vector<format_t*>& xf_formats);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_style_sheet() const;
  [[nodiscard]] std::string write_num_fmts() const;
  [[nodiscard]] std::string write_font_condense() const;
  [[nodiscard]] std::string write_font_extend() const;
  [[nodiscard]] std::string write_font_underline(format_underlines_t underline) const;
  [[nodiscard]] std::string write_font_vert_align(const std::string& align) const;
  [[nodiscard]] std::string write_font_size(double font_size) const;
  [[nodiscard]] std::string write_font_color_theme(uint8_t theme) const;
  [[nodiscard]] std::string write_font_color_indexed(uint8_t index) const;
  [[nodiscard]] std::string write_font_color_rgb(color_t rgb) const;
  [[nodiscard]] std::string write_font_name(const std::string& font_name, bool is_rich_string) const;
  [[nodiscard]] std::string write_font_family(uint8_t font_family) const;
  [[nodiscard]] std::string write_font_charset(uint8_t font_charset) const;
  [[nodiscard]] std::string write_font_scheme(const std::string& font_scheme) const;
  [[nodiscard]] std::string write_font(const format_t* format, bool is_dxf, bool is_rich_string) const;
  [[nodiscard]] std::string write_fonts() const;
  [[nodiscard]] std::string write_default_fill(const std::string& pattern) const;
  [[nodiscard]] std::string write_fills() const;
  [[nodiscard]] std::string convert_format_borders_style(format_borders_t style) const;
  [[nodiscard]] std::string write_border_color(color_t color) const;
  [[nodiscard]] std::string write_sub_border(const std::string& type, format_borders_t style, color_t color) const;
  [[nodiscard]] std::string write_border(const format_t* format, bool is_dxf) const;
  [[nodiscard]] std::string write_borders() const;
  [[nodiscard]] std::string write_style_xf(bool has_hyperlink, uint16_t font_id) const;
  [[nodiscard]] std::string write_cell_style_xfs() const;
  [[nodiscard]] std::string write_cell_xfs() const;
  [[nodiscard]] std::string write_cell_style(const std::string& name, uint8_t xf_id, uint8_t builtin_id) const;
  [[nodiscard]] std::string write_cell_styles() const;
  [[nodiscard]] std::string write_dxfs() const;
  [[nodiscard]] std::string write_table_styles() const;

  uint32_t font_count_;
  ///  uint32_t xf_count;
  ///  uint32_t dxf_count;
  ///  uint32_t num_format_count;
  uint32_t border_count_;
  ///  uint32_t fill_count;
  std::vector<format_t*> xf_formats_;
  std::vector<format_t*> dxf_formats_;
  ///  uint8_t has_hyperlink;
  ///  uint16_t hyperlink_font_id;
  ///  uint8_t has_comments;
};

/// void lxw_styles_write_string_fragment(lxw_styles *self, const char *string);
/// void lxw_styles_write_rich_font(lxw_styles *styles, lxw_format *format);

/// STATIC void _write_xf(lxw_styles *self, lxw_format *format);
}

#endif
