/*
 * format - A library for creating Excel XLSX format files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/format.h"

#include "xwpp/exception.h"

#include <format>

namespace xwpp
{

format_t::format_t(std::function<int32_t(format_t*)> get_dxf_index)
  : get_dxf_index_{std::move(get_dxf_index)}
{
}

void format_t::set_num_format(const std::string& num_format)
{
  num_format_ = num_format;
}

void format_t::set_bold()
{
  bold_ = true;
}

void format_t::set_italic()
{
  italic_ = true;
}

void format_t::set_align(format_alignments_t alignment)
{
  switch(alignment)
  {
    case format_alignments_t::HORIZONTAL_LEFT:
    case format_alignments_t::HORIZONTAL_CENTER:
    case format_alignments_t::HORIZONTAL_RIGHT:
    case format_alignments_t::HORIZONTAL_FILL:
    case format_alignments_t::HORIZONTAL_JUSTIFY:
    case format_alignments_t::HORIZONTAL_CENTER_ACROSS:
    case format_alignments_t::HORIZONTAL_DISTRIBUTED:
      text_h_align_ = alignment;
      break;

    case format_alignments_t::VERTICAL_TOP:
    case format_alignments_t::VERTICAL_BOTTOM:
    case format_alignments_t::VERTICAL_CENTER:
    case format_alignments_t::VERTICAL_JUSTIFY:
    case format_alignments_t::VERTICAL_DISTRIBUTED:
      text_v_align_ = alignment;
      break;

    case format_alignments_t::NONE:
      // NOP
      break;
  }
}

void format_t::set_font_color(color_t color)
{
  font_color_ = color;
}

void format_t::set_fg_color(color_t color)
{
  fg_color_ = color;
}

void format_t::set_bg_color(color_t color)
{
  bg_color_ = color;
}

void format_t::set_underline(format_underlines_t style)
{
  switch(style)
  {
    case format_underlines_t::SINGLE:
    case format_underlines_t::DOUBLE:
    case format_underlines_t::SINGLE_ACCOUNTING:
    case format_underlines_t::DOUBLE_ACCOUNTING:
      underline_ = style;
      break;

    case format_underlines_t::NONE:
      // NOP
      break;
  }
}

void format_t::set_text_wrap()
{
  text_wrap_ = true;
}

void format_t::set_border(format_borders_t style)
{
  style = check_border(style);

  bottom_ = style;
  top_    = style;
  left_   = style;
  right_  = style;
}

void format_t::set_bottom(format_borders_t style)
{
  bottom_ = check_border(style);
}

void format_t::set_top(format_borders_t style)
{
  top_ = check_border(style);
}

void format_t::set_left(format_borders_t style)
{
  left_ = check_border(style);
}

void format_t::set_right(format_borders_t style)
{
  right_ = check_border(style);
}

void format_t::set_pattern(format_patterns_t pattern)
{
  pattern_ = pattern;
}

void format_t::set_border_color(color_t color)
{
  bottom_color_ = color;
  top_color_    = color;
  left_color_   = color;
  right_color_  = color;
}

void format_t::set_bottom_color(color_t color)
{
  bottom_color_ = color;
}

void format_t::set_top_color(color_t color)
{
  top_color_ = color;
}

void format_t::set_left_color(color_t color)
{
  left_color_ = color;
}

void format_t::set_right_color(color_t color)
{
  right_color_ = color;
}

void format_t::set_diag_type(format_diagonal_types_t type)
{
  diag_type_ = type;
}

void format_t::set_diag_color(color_t color)
{
  diag_color_ = color;
}

void format_t::set_diag_border(format_borders_t style)
{
  diag_border_ = style;
}

void format_t::set_num_format_index(uint8_t index)
{
  num_format_index_ = index;
}

void format_t::set_font_name(const std::string& font_name)
{
  font_name_ = font_name;
}

void format_t::set_font_size(double size)
{
  if(size >= MIN_FONT_SIZE && size <= MAX_FONT_SIZE)
  {
    font_size_ = size;
  }
}

void format_t::set_font_strikeout()
{
  font_strikeout_ = true;
}

void format_t::set_font_script(format_scripts_t style)
{
  font_script_ = style;
}

void format_t::set_font_family(uint8_t value)
{
  font_family_ = value;
}

void format_t::set_font_charset(uint8_t value)
{
  font_charset_ = value;
}

void format_t::set_font_outline()
{
  font_outline_ = true;
}

void format_t::set_font_shadow()
{
  font_shadow_ = true;
}

void format_t::set_font_scheme(const std::string& font_scheme)
{
  font_scheme_ = font_scheme;
}

void format_t::set_font_condense()
{
  font_condense_ = true;
}

void format_t::set_font_extend()
{
  font_extend_ = true;
}

void format_t::set_font_only()
{
  font_only_ = true;
}

void format_t::set_unlocked()
{
  locked_ = false;
}

void format_t::set_hidden()
{
  hidden_ = true;
}

void format_t::set_rotation(int16_t angle)
{
  // Convert user angle to Excel angle.
  if(angle == 270)
  {
    rotation_ = 255;
  }
  else if(angle >= -90 && angle <= 90)
  {
    if(angle < 0)
    {
      // NOLINTNEXTLINE(bugprone-narrowing-conversions,cppcoreguidelines-narrowing-conversions)
      angle = -angle + 90;
    }

    rotation_ = angle;
  }
  else
  {
    throw xwpp_out_of_range_t(
      std::format("format_t::set_rotation(): format rotation '{}' outside Excel range: -90 <= rotation <= 90", angle));
  }
}

void format_t::set_indent(uint8_t level)
{
  indent_ = level;
}

void format_t::set_shrink()
{
  shrink_ = true;
}

void format_t::set_text_justlast()
{
  text_justlast_ = true;
}

void format_t::set_valign(format_alignments_t alignment)
{
  if(alignment == format_alignments_t::VERTICAL_TOP || alignment == format_alignments_t::VERTICAL_BOTTOM ||
     alignment == format_alignments_t::VERTICAL_CENTER || alignment == format_alignments_t::VERTICAL_JUSTIFY ||
     alignment == format_alignments_t::VERTICAL_DISTRIBUTED)
  {
    text_v_align_ = alignment;
  }
  else
  {
    throw xwpp_exception_t(
      std::format("format_t::set_valign(): alignment '{}' is not a vertical alignement", static_cast<int>(alignment)));
  }
}

void format_t::set_halign(format_alignments_t alignment)
{
  if(alignment == format_alignments_t::HORIZONTAL_LEFT || alignment == format_alignments_t::HORIZONTAL_CENTER ||
     alignment == format_alignments_t::HORIZONTAL_RIGHT || alignment == format_alignments_t::HORIZONTAL_FILL ||
     alignment == format_alignments_t::HORIZONTAL_JUSTIFY ||
     alignment == format_alignments_t::HORIZONTAL_CENTER_ACROSS ||
     alignment == format_alignments_t::HORIZONTAL_DISTRIBUTED)
  {
    text_h_align_ = alignment;
  }
  else
  {
    throw xwpp_exception_t(std::format("format_t::set_halign(): alignment '{}' is not a horizontal alignement",
                                       static_cast<int>(alignment)));
  }
}

void format_t::set_reading_order(uint8_t value)
{
  reading_order_ = value;
}

void format_t::set_theme(uint8_t value)
{
  theme_ = value;
}

void format_t::set_color_indexed(uint8_t value)
{
  color_indexed_ = value;
}

void format_t::set_quote_prefix()
{
  quote_prefix_ = true;
}

void format_t::set_hyperlink()
{
  hyperlink_ = true;
  xf_id_     = 1;
  underline_ = format_underlines_t::SINGLE;
  theme_     = 10;
}

format_borders_t format_t::check_border(format_borders_t style) const
{
  switch(style)
  {
    case format_borders_t::NONE:
    case format_borders_t::THIN:
    case format_borders_t::MEDIUM:
    case format_borders_t::DASHED:
    case format_borders_t::DOTTED:
    case format_borders_t::THICK:
    case format_borders_t::DOUBLE:
    case format_borders_t::HAIR:
    case format_borders_t::MEDIUM_DASHED:
    case format_borders_t::DASH_DOT:
    case format_borders_t::MEDIUM_DASH_DOT:
    case format_borders_t::DASH_DOT_DOT:
    case format_borders_t::MEDIUM_DASH_DOT_DOT:
    case format_borders_t::SLANT_DASH_DOT:
      return style;

    default:
      return format_borders_t::NONE;
  }
}

const std::string format_t::DEFAULT_FONT_NAME = "Calibri";

}
