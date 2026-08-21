/*
 * styles - A Xlsxwriter++ library for creating Excel XLSX styles files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
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
  style_t() = default;
  style_t(uint32_t font_count, uint32_t fill_count, uint32_t border_count, uint32_t num_format_count, bool has_comments,
          const std::vector<format_t*>& xf_formats, const std::vector<format_t*>& dxf_formats);

  [[nodiscard]] std::string assemble_xml_file();
  [[nodiscard]] static std::string write_string_fragment(const std::string& str);
  [[nodiscard]] std::string write_rich_font(const format_t* format);

private:
  [[nodiscard]] std::string write_num_fmts() const;
  [[nodiscard]] std::string write_font(const format_t* format, bool is_dxf, bool is_rich_string);
  [[nodiscard]] std::string write_fonts();
  [[nodiscard]] std::string write_fills() const;
  [[nodiscard]] static std::string write_border(const format_t* format, bool is_dxf);
  [[nodiscard]] std::string write_borders() const;
  [[nodiscard]] std::string write_cell_style_xfs() const;
  [[nodiscard]] std::string write_cell_xfs() const;
  [[nodiscard]] std::string write_cell_styles() const;
  [[nodiscard]] static bool apply_alignment(const format_t* format);
  [[nodiscard]] static bool has_alignment(const format_t* format);
  [[nodiscard]] static std::string write_alignment(const format_t* format);
  [[nodiscard]] static std::string write_protection(const format_t* format);
  [[nodiscard]] static std::string write_xf(const format_t* format);
  [[nodiscard]] std::string write_dxfs();
  [[nodiscard]] static std::string write_fill(const format_t* format, bool is_dxf);

  uint32_t font_count_       = 0;
  uint32_t num_format_count_ = 0;
  uint32_t border_count_     = 0;
  uint32_t fill_count_       = 0;
  std::vector<format_t*> xf_formats_;
  std::vector<format_t*> dxf_formats_;
  bool has_hyperlink_        = false;
  int32_t hyperlink_font_id_ = 0;
  bool has_comments_         = false;
};

}

#endif
