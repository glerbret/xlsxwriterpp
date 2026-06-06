/*
 * styles - A library for creating Excel XLSX styles files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/styles.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <cassert>
#include <format>
#include <string>

namespace xwpp
{

/// void lxw_styles_write_string_fragment(lxw_styles *self, const char *string)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

/* Add attribute to preserve leading or trailing whitespace. */
///     if (isspace((unsigned char) string[0])
///         || isspace((unsigned char) string[strlen(string) - 1]))
///         LXW_PUSH_ATTRIBUTES_STR("xml:space", "preserve");

///     lxw_xml_data_element(self->file, "t", string, &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// void
/// lxw_styles_write_rich_font(lxw_styles *self, lxw_format *format)
/// {
///     _write_font(self, format, LXW_FALSE, LXW_TRUE);
/// }

std::string style_t::write_style_sheet() const
{
  return xml_start_tag("styleSheet", {
                                         {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
  });
}

/// STATIC void
/// _write_num_fmt(lxw_styles *self, uint16_t num_fmt_id, char *format_code)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char *format_codes[] = {
///         "General",
///         "0",
///         "0.00",
///         "#,##0",
///         "#,##0.00",
///         "($#,##0_);($#,##0)",
///         "($#,##0_);[Red]($#,##0)",
///         "($#,##0.00_);($#,##0.00)",
///         "($#,##0.00_);[Red]($#,##0.00)",
///         "0%",
///         "0.00%",
///         "0.00E+00",
///         "# ?/?",
///         "# ?" "?/?" "?",        /* Split string to avoid unintentional trigraph. */
///         "m/d/yy",
///         "d-mmm-yy",
///         "d-mmm",
///         "mmm-yy",
///         "h:mm AM/PM",
///         "h:mm:ss AM/PM",
///         "h:mm",
///         "h:mm:ss",
///         "m/d/yy h:mm",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "General",
///         "(#,##0_);(#,##0)",
///         "(#,##0_);[Red](#,##0)",
///         "(#,##0.00_);(#,##0.00)",
///         "(#,##0.00_);[Red](#,##0.00)",
///         "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",
///         "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)",
///         "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",
///         "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
///         "mm:ss",
///         "[h]:mm:ss",
///         "mm:ss.0",
///         "##0.0E+0",
///         "@"
///     };

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("numFmtId", num_fmt_id);

///     if (num_fmt_id < 50)
///         LXW_PUSH_ATTRIBUTES_STR("formatCode", format_codes[num_fmt_id]);
///     else if (num_fmt_id < 164)
///         LXW_PUSH_ATTRIBUTES_STR("formatCode", "General");
///     else
///         LXW_PUSH_ATTRIBUTES_STR("formatCode", format_code);

///     lxw_xml_empty_tag(self->file, "numFmt", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

std::string style_t::write_num_fmts() const
{
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     lxw_format *format;
  ///     uint16_t last_format_index = 0;

  ///     if (!self->num_format_count)
  ///         return;

  ///     LXW_INIT_ATTRIBUTES();
  ///     LXW_PUSH_ATTRIBUTES_INT("count", self->num_format_count);

  ///     lxw_xml_start_tag(self->file, "numFmts", &attributes);

  /* Write the numFmts elements. */
  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {

  /* Ignore built-in number formats, i.e., < 164. */
  ///         if (format->num_format_index < 164)
  ///             continue;

  /* Ignore duplicates which have an already used index. */
  ///         if (format->num_format_index <= last_format_index)
  ///             continue;

  ///         _write_num_fmt(self, format->num_format_index, format->num_format);

  ///         last_format_index = format->num_format_index;
  ///     }

  ///     lxw_xml_end_tag(self->file, "numFmts");

  ///     LXW_FREE_ATTRIBUTES();
  return "";
}

std::string style_t::write_font_size(double font_size) const
{
  return xml_empty_tag("sz", {
                                 {"val", std::format("{}", font_size)}
  });
}

std::string style_t::write_font_color_theme(uint8_t theme) const
{
  return xml_empty_tag("color", {
                                    {"theme", std::format("{:d}", theme)}
  });
}

std::string style_t::write_font_color_rgb(color_t rgb) const
{
  return xml_empty_tag("color", {
                                    {"rgb", std::to_string(static_cast<uint32_t>(rgb))}
  });
}

std::string style_t::write_font_color_indexed(uint8_t index) const
{
  return xml_empty_tag("color", {
                                    {"indexed", std::format("{:d}", index)}
  });
}

std::string style_t::write_font_name(const std::string& font_name, bool is_rich_string) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(!font_name.empty())
  {
    attributes.emplace_back("val", font_name);
  }
  else
  {
    attributes.emplace_back("val", format_t::DEFAULT_FONT_NAME);
  }

  if(is_rich_string)
  {
    return xml_empty_tag("rFont", attributes);
  }
  else
  {
    return xml_empty_tag("name", attributes);
  }
}

std::string style_t::write_font_family(uint8_t font_family) const
{
  return xml_empty_tag("family", {
                                     {"val", std::format("{:d}", font_family)}
  });
}

std::string style_t::write_font_charset(uint8_t font_charset) const
{
  return xml_empty_tag("charset", {
                                      {"val", std::format("{:d}", font_charset)}
  });
}

std::string style_t::write_font_scheme(const std::string& font_scheme) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  if(!font_scheme.empty())
  {
    attributes.emplace_back("val", font_scheme);
  }
  else
  {
    attributes.emplace_back("val", "minor");
  }

  return xml_empty_tag("scheme", attributes);
}

std::string style_t::write_font_underline(format_underlines_t underline) const
{
  switch(underline)
  {
    case format_underlines_t::SINGLE:
      return xml_empty_tag("u");

    case format_underlines_t::DOUBLE:
      return xml_empty_tag("u", {
                                    {"val", "double"}
      });

    case format_underlines_t::SINGLE_ACCOUNTING:
      return xml_empty_tag("u", {
                                    {"val", "singleAccounting"}
      });

    case format_underlines_t::DOUBLE_ACCOUNTING:
      return xml_empty_tag("u", {
                                    {"val", "doubleAccounting"}
      });

    case format_underlines_t::NONE:
    default:
      return "";
  }
}

std::string style_t::write_font_condense() const
{
  return xml_empty_tag("condense", {
                                       {"val", "0"}
  });
}

std::string style_t::write_font_extend() const
{
  return xml_empty_tag("extend", {
                                     {"val", "0"}
  });
}

std::string style_t::write_font_vert_align(const std::string& align) const
{
  return xml_empty_tag("vertAlign", {
                                        {"val", align}
  });
}

[[nodiscard]] std::string style_t::write_font(const format_t* format, bool is_dxf, bool is_rich_string) const
{
  std::string xml_data;

  if(is_rich_string)
  {
    xml_data += xml_start_tag("rPr");
  }
  else
  {
    xml_data += xml_start_tag("font");
  }

  if(format->font_condense_)
  {
    xml_data += write_font_condense();
  }

  if(format->font_extend_)
  {
    xml_data += write_font_extend();
  }

  if(format->bold_)
  {
    xml_data += xml_empty_tag("b");
  }

  if(format->italic_)
  {
    xml_data += xml_empty_tag("i");
  }

  if(format->font_strikeout_)
  {
    xml_data += xml_empty_tag("strike");
  }

  if(format->font_outline_)
  {
    xml_data += xml_empty_tag("outline");
  }

  if(format->font_shadow_)
  {
    xml_data += xml_empty_tag("shadow");
  }

  if(format->underline_ != format_underlines_t::NONE)
  {
    xml_data += write_font_underline(format->underline_);
  }

  if(format->font_script_ == format_scripts_t::SUPERSCRIPT)
  {
    xml_data += write_font_vert_align("superscript");
  }

  if(format->font_script_ == format_scripts_t::SUBSCRIPT)
  {
    xml_data += write_font_vert_align("subscript");
  }

  if(!is_dxf && format->font_size_ > 0.0)
  {
    xml_data += write_font_size(format->font_size_);
  }

  if(format->theme_)
  {
    xml_data += write_font_color_theme(format->theme_);
  }
  else if(format->color_indexed_)
  {
    xml_data += write_font_color_indexed(format->color_indexed_);
  }
  else if(format->font_color_ != color_t::UNSET)
  {
    xml_data += write_font_color_rgb(format->font_color_);
  }
  else if(!is_dxf)
  {
    // Default font theme
    xml_data += write_font_color_theme(1);
  }

  if(!is_dxf)
  {
    xml_data += write_font_name(format->font_name_, is_rich_string);
    xml_data += write_font_family(format->font_family_);

    if(format->font_charset_)
    {
      xml_data += write_font_charset(format->font_charset_);
    }

    // Only write the scheme element for the default font type if it isn't a hyperlink.
    if((format->font_name_.empty() || format->font_name_ == format_t::DEFAULT_FONT_NAME) && !format->hyperlink_)
    {
      xml_data += write_font_scheme(format->font_scheme_);
    }
  }

  ///     if (format->hyperlink) {
  ///         self->has_hyperlink = LXW_TRUE;

  ///         if (self->hyperlink_font_id == 0)
  ///             self->hyperlink_font_id = format->font_index;
  ///     }

  if(is_rich_string)
  {
    xml_data += xml_end_tag("rPr");
  }
  else
  {
    xml_data += xml_end_tag("font");
  }

  return xml_data;
}

/// STATIC void
/// _write_comment_font(lxw_styles *self)
/// {
///     lxw_xml_start_tag(self->file, "font", NULL);

///     _write_font_size(self, 8);
///     _write_font_color_indexed(self, 81);
///     _write_font_name(self, "Tahoma", LXW_FALSE);
///     _write_font_family(self, 2);

///     lxw_xml_end_tag(self->file, "font");
/// }

style_t::style_t(uint32_t font_count, uint32_t border_count, const std::vector<format_t*>& xf_formats)
  : font_count_{font_count}
  , border_count_{border_count}
  , xf_formats_{xf_formats}
{
}

std::string style_t::write_fonts() const
{
  ///    struct xml_attribute_list attributes;
  ///    struct xml_attribute *attribute;
  ///    lxw_format *format;
  const uint16_t count = font_count_;
  ///    if (self->has_comments)
  ///        count++;

  std::string xml_data = xml_start_tag("fonts", {
                                                    {"count", std::to_string(count)}
  });

  for(const auto format: xf_formats_)
  {
    if(format->has_font_)
    {
      xml_data += write_font(format, false, false);
    }
  }

  ///     if (self->has_comments)
  ///         _write_comment_font(self);

  xml_data += xml_end_tag("fonts");

  return xml_data;
}

std::string style_t::write_default_fill(const std::string& pattern) const
{
  std::string xml_data = xml_start_tag("fill");
  xml_data += xml_empty_tag("patternFill", {
                                               {"patternType", pattern}
  });
  xml_data += xml_end_tag("fill");

  return xml_data;
}

/// STATIC void
/// _write_fg_color(lxw_styles *self, lxw_color_t color)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb_str[LXW_ATTR_32];

///     LXW_INIT_ATTRIBUTES();

///     lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
///     LXW_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

///     lxw_xml_empty_tag(self->file, "fgColor", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_bg_color(lxw_styles *self, lxw_color_t color, uint8_t pattern)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb_str[LXW_ATTR_32];

///     LXW_INIT_ATTRIBUTES();

///     if (color == LXW_COLOR_UNSET) {
///         if (pattern <= LXW_PATTERN_SOLID) {
///             LXW_PUSH_ATTRIBUTES_STR("indexed", "64");
///             lxw_xml_empty_tag(self->file, "bgColor", &attributes);
///         }
///     }
///     else {
///         lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
///         LXW_PUSH_ATTRIBUTES_STR("rgb", rgb_str);
///         lxw_xml_empty_tag(self->file, "bgColor", &attributes);
///     }

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_fill(lxw_styles *self, lxw_format *format, uint8_t is_dxf)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     uint8_t pattern = format->pattern;
///     lxw_color_t bg_color = format->bg_color;
///     lxw_color_t fg_color = format->fg_color;

///     char *patterns[] = {
///         "none",
///         "solid",
///         "mediumGray",
///         "darkGray",
///         "lightGray",
///         "darkHorizontal",
///         "darkVertical",
///         "darkDown",
///         "darkUp",
///         "darkGrid",
///         "darkTrellis",
///         "lightHorizontal",
///         "lightVertical",
///         "lightDown",
///         "lightUp",
///         "lightGrid",
///         "lightTrellis",
///         "gray125",
///         "gray0625",
///     };

///     if (is_dxf) {
///         bg_color = format->dxf_bg_color;
///         fg_color = format->dxf_fg_color;
///     }

///     LXW_INIT_ATTRIBUTES();

/* Special handling for pattern only case. */
///     if (!bg_color && !fg_color && pattern) {
///         _write_default_fill(self, patterns[pattern]);
///         LXW_FREE_ATTRIBUTES();
///         return;
///     }

///     lxw_xml_start_tag(self->file, "fill", NULL);

/* None/Solid patterns are handled differently for dxf formats. */
///     if (pattern && !(is_dxf && pattern <= LXW_PATTERN_SOLID))
///         LXW_PUSH_ATTRIBUTES_STR("patternType", patterns[pattern]);

///     lxw_xml_start_tag(self->file, "patternFill", &attributes);

///     if (fg_color != LXW_COLOR_UNSET)
///         _write_fg_color(self, fg_color);

///     _write_bg_color(self, bg_color, pattern);

///     lxw_xml_end_tag(self->file, "patternFill");
///     lxw_xml_end_tag(self->file, "fill");

///     LXW_FREE_ATTRIBUTES();
/// }

std::string style_t::write_fills() const
{
  ///     lxw_format *format;

  std::string xml_data = xml_start_tag("fills", {
                                                    {"count", "2" /* TODO self->fill_count*/}
  });

  // Write the default fills.
  xml_data += write_default_fill("none");
  xml_data += write_default_fill("gray125");

  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
  ///         if (format->has_fill)
  ///             _write_fill(self, format, LXW_FALSE);
  ///     }

  xml_data += xml_end_tag("fills");

  return xml_data;
}

std::string style_t::write_border_color(color_t color) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(color != color_t::UNSET)
  {
    attributes.emplace_back("rgb", std::format("FF{:06X}", static_cast<uint32_t>(color) & COLOR_MASK));
  }
  else
  {
    attributes.emplace_back("auto", "1");
  }

  return xml_empty_tag("color", attributes);
}

std::string style_t::convert_format_borders_style(format_borders_t style) const
{
  switch(style)
  {
    case format_borders_t::THIN:
      return "thin";

    case format_borders_t::MEDIUM:
      return "medium";

    case format_borders_t::DASHED:
      return "dashed";

    case format_borders_t::DOTTED:
      return "dotted";

    case format_borders_t::THICK:
      return "thick";

    case format_borders_t::DOUBLE:
      return "double";

    case format_borders_t::HAIR:
      return "hair";

    case format_borders_t::MEDIUM_DASHED:
      return "mediumDashed";

    case format_borders_t::DASH_DOT:
      return "dashDot";

    case format_borders_t::MEDIUM_DASH_DOT:
      return "mediumDashDot";

    case format_borders_t::DASH_DOT_DOT:
      return "dashDotDot";

    case format_borders_t::MEDIUM_DASH_DOT_DOT:
      return "mediumDashDotDot";

    case format_borders_t::SLANT_DASH_DOT:
      return "slantDashDot";

    case format_borders_t::NONE:
    default:
      return "none";
  }
}

std::string style_t::write_sub_border(const std::string& type, format_borders_t style, color_t color) const
{
  if(style == format_borders_t::NONE)
  {
    return xml_empty_tag(type);
  }

  std::string xml_data = xml_start_tag(type, {
                                                 {"style", convert_format_borders_style(style)}
  });
  xml_data += write_border_color(color);
  xml_data += xml_end_tag(type);

  return xml_data;
}

std::string style_t::write_border(const format_t* format, bool is_dxf) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  // Add attributes for diagonal borders.
  if(format->diag_type_ == format_diagonal_types_t::BORDER_UP)
  {
    attributes.emplace_back("diagonalUp", "1");
  }
  else if(format->diag_type_ == format_diagonal_types_t::BORDER_DOWN)
  {
    attributes.emplace_back("diagonalDown", "1");
  }
  else if(format->diag_type_ == format_diagonal_types_t::BORDER_UP_DOWN)
  {
    attributes.emplace_back("diagonalUp", "1");
    attributes.emplace_back("diagonalDown", "1");
  }

  // Ensure that a default diag border is set if the diag type is set.
  format_borders_t diag_border = format->diag_border_;
  if(format->diag_type_ != format_diagonal_types_t::NONE && diag_border == format_borders_t::NONE)
  {
    diag_border = format_borders_t::THIN;
  }

  // Write the start border tag.
  std::string xml_data = xml_start_tag("border", attributes);

  // Write the <border> sub elements.
  xml_data += write_sub_border("left", format->left_, format->left_color_);
  xml_data += write_sub_border("right", format->right_, format->right_color_);
  xml_data += write_sub_border("top", format->top_, format->top_color_);
  xml_data += write_sub_border("bottom", format->bottom_, format->bottom_color_);

  if(is_dxf)
  {
    xml_data += write_sub_border("vertical", format_borders_t::NONE, color_t::UNSET);
    xml_data += write_sub_border("horizontal", format_borders_t::NONE, color_t::UNSET);
  }

  // Conditional DXF formats don't allow diagonal borders.
  if(!is_dxf)
  {
    xml_data += write_sub_border("diagonal", diag_border, format->diag_color_);
  }

  xml_data += xml_end_tag("border");

  return xml_data;
}

std::string style_t::write_borders() const
{
  std::string xml_data = xml_start_tag("borders", {
                                                      {"count", std::to_string(border_count_)}
  });

  for(const auto format: xf_formats_)
  {
    if(format->has_border_)
    {
      xml_data += write_border(format, false);
    }
  }

  xml_data += xml_end_tag("borders");

  return xml_data;
}

/// STATIC void
/// _write_hyperlink_alignment(lxw_styles *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("vertical", "top");

///     lxw_xml_empty_tag(self->file, "alignment", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_hyperlink_protection(lxw_styles *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("locked", "0");

///     lxw_xml_empty_tag(self->file, "protection", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

std::string style_t::write_style_xf(bool has_hyperlink, uint16_t font_id) const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"numFmtId", "0"                    },
      {"fontId",   std::to_string(font_id)},
      {"fillId",   "0"                    },
      {"borderId", "0"                    },
  };

  if(has_hyperlink)
  {
    attributes.emplace_back("applyNumberFormat", "0");
    attributes.emplace_back("applyFill", "0");
    attributes.emplace_back("applyBorder", "0");
    attributes.emplace_back("applyAlignment", "0");
    attributes.emplace_back("applyProtection", "0");

    std::string xml_data = xml_start_tag("xf", attributes);
    ///         _write_hyperlink_alignment(self);
    ///        _write_hyperlink_protection(self);
    xml_data += xml_end_tag("xf");
    return xml_data;
  }
  else
  {
    return xml_empty_tag("xf", attributes);
  }
}

std::string style_t::write_cell_style_xfs() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  ///   if (self->has_hyperlink)
  ///     LXW_PUSH_ATTRIBUTES_STR("count", "2");
  ///   else
  attributes.emplace_back("count", "1");

  std::string xml_data = xml_start_tag("cellStyleXfs", attributes);
  xml_data += write_style_xf(false, 0);

  ///   if (self->has_hyperlink)
  ///     _write_style_xf(self, self->has_hyperlink, self->hyperlink_font_id);

  xml_data += xml_end_tag("cellStyleXfs");

  return xml_data;
}

/// STATIC uint8_t
/// _apply_alignment(lxw_format *format)
/// {
///     return format->text_h_align != LXW_ALIGN_NONE
///         || format->text_v_align != LXW_ALIGN_NONE
///         || format->indent != 0
///         || format->rotation != 0
///         || format->text_wrap != 0
///         || format->shrink != 0 || format->reading_order != 0;
/// }

/// STATIC uint8_t
/// _has_alignment(lxw_format *format)
/// {
///     return format->text_h_align != LXW_ALIGN_NONE
///         || !(format->text_v_align == LXW_ALIGN_NONE ||
///              format->text_v_align == LXW_ALIGN_VERTICAL_BOTTOM)
///         || format->indent != 0
///         || format->rotation != 0
///         || format->text_wrap != 0
///         || format->shrink != 0 || format->reading_order != 0;
/// }

/// STATIC void
/// _write_alignment(lxw_styles *self, lxw_format *format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     int16_t rotation = format->rotation;

///     LXW_INIT_ATTRIBUTES();

/* Indent is only allowed for some alignment properties. */
/* If it is defined for any other alignment or no alignment has been  */
/* set then default to left alignment. */
///     if (format->indent
///         && format->text_h_align != LXW_ALIGN_LEFT
///         && format->text_h_align != LXW_ALIGN_RIGHT
///         && format->text_h_align != LXW_ALIGN_DISTRIBUTED
///         && format->text_v_align != LXW_ALIGN_VERTICAL_TOP
///         && format->text_v_align != LXW_ALIGN_VERTICAL_BOTTOM
///         && format->text_v_align != LXW_ALIGN_VERTICAL_DISTRIBUTED) {
///         format->text_h_align = LXW_ALIGN_LEFT;
///     }

/* Check for properties that are mutually exclusive. */
///     if (format->text_wrap)
///         format->shrink = 0;

///     if (format->text_h_align == LXW_ALIGN_FILL)
///         format->shrink = 0;

///     if (format->text_h_align == LXW_ALIGN_JUSTIFY)
///         format->shrink = 0;

///     if (format->text_h_align == LXW_ALIGN_DISTRIBUTED)
///         format->shrink = 0;

///     if (format->text_h_align != LXW_ALIGN_DISTRIBUTED)
///         format->just_distrib = 0;

///     if (format->indent)
///         format->just_distrib = 0;

///     if (format->text_h_align == LXW_ALIGN_LEFT)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "left");

///     if (format->text_h_align == LXW_ALIGN_CENTER)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "center");

///     if (format->text_h_align == LXW_ALIGN_RIGHT)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "right");

///     if (format->text_h_align == LXW_ALIGN_FILL)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "fill");

///     if (format->text_h_align == LXW_ALIGN_JUSTIFY)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "justify");

///     if (format->text_h_align == LXW_ALIGN_CENTER_ACROSS)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "centerContinuous");

///     if (format->text_h_align == LXW_ALIGN_DISTRIBUTED)
///         LXW_PUSH_ATTRIBUTES_STR("horizontal", "distributed");

///     if (format->just_distrib)
///         LXW_PUSH_ATTRIBUTES_STR("justifyLastLine", "1");

///     if (format->text_v_align == LXW_ALIGN_VERTICAL_TOP)
///         LXW_PUSH_ATTRIBUTES_STR("vertical", "top");

///     if (format->text_v_align == LXW_ALIGN_VERTICAL_CENTER)
///         LXW_PUSH_ATTRIBUTES_STR("vertical", "center");

///     if (format->text_v_align == LXW_ALIGN_VERTICAL_JUSTIFY)
///         LXW_PUSH_ATTRIBUTES_STR("vertical", "justify");

///     if (format->text_v_align == LXW_ALIGN_VERTICAL_DISTRIBUTED)
///         LXW_PUSH_ATTRIBUTES_STR("vertical", "distributed");

/* Map rotation to Excel values. */
///     if (rotation) {
///         if (rotation == 270)
///             rotation = 255;
///         else if (rotation < 0)
///             rotation = -rotation + 90;

///         LXW_PUSH_ATTRIBUTES_INT("textRotation", rotation);
///     }

///     if (format->indent)
///         LXW_PUSH_ATTRIBUTES_INT("indent", format->indent);

///     if (format->text_wrap)
///         LXW_PUSH_ATTRIBUTES_STR("wrapText", "1");

///     if (format->shrink)
///         LXW_PUSH_ATTRIBUTES_STR("shrinkToFit", "1");

///     if (format->reading_order == 1)
///         LXW_PUSH_ATTRIBUTES_STR("readingOrder", "1");

///     if (format->reading_order == 2)
///         LXW_PUSH_ATTRIBUTES_STR("readingOrder", "2");

///     if (!STAILQ_EMPTY(&attributes))
///         lxw_xml_empty_tag(self->file, "alignment", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_protection(lxw_styles *self, lxw_format *format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (!format->locked)
///         LXW_PUSH_ATTRIBUTES_STR("locked", "0");

///     if (format->hidden)
///         LXW_PUSH_ATTRIBUTES_STR("hidden", "1");

///     lxw_xml_empty_tag(self->file, "protection", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_xf(lxw_styles *self, lxw_format *format)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     uint8_t has_protection = (!format->locked) | format->hidden;
///     uint8_t has_alignment = _has_alignment(format);
///     uint8_t apply_alignment = _apply_alignment(format);

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("numFmtId", format->num_format_index);
///     LXW_PUSH_ATTRIBUTES_INT("fontId", format->font_index);
///     LXW_PUSH_ATTRIBUTES_INT("fillId", format->fill_index);
///     LXW_PUSH_ATTRIBUTES_INT("borderId", format->border_index);
///     LXW_PUSH_ATTRIBUTES_INT("xfId", format->xf_id);

///     if (format->quote_prefix)
///         LXW_PUSH_ATTRIBUTES_STR("quotePrefix", "1");

///     if (format->num_format_index > 0)
///         LXW_PUSH_ATTRIBUTES_STR("applyNumberFormat", "1");

/* Add applyFont attribute if XF format uses a font element. */
///     if (format->font_index > 0 && !format->hyperlink)
///         LXW_PUSH_ATTRIBUTES_STR("applyFont", "1");

/* Add applyFill attribute if XF format uses a fill element. */
///     if (format->fill_index > 0)
///         LXW_PUSH_ATTRIBUTES_STR("applyFill", "1");

/* Add applyBorder attribute if XF format uses a border element. */
///     if (format->border_index > 0)
///         LXW_PUSH_ATTRIBUTES_STR("applyBorder", "1");

/* We can also have applyAlignment without a sub-element. */
///     if (apply_alignment || format->hyperlink)
///         LXW_PUSH_ATTRIBUTES_STR("applyAlignment", "1");

///     if (has_protection || format->hyperlink)
///         LXW_PUSH_ATTRIBUTES_STR("applyProtection", "1");

/* Write XF with sub-elements if required. */
///     if (has_alignment || has_protection) {
///         lxw_xml_start_tag(self->file, "xf", &attributes);

///         if (has_alignment)
///             _write_alignment(self, format);

///         if (has_protection)
///             _write_protection(self, format);

///         lxw_xml_end_tag(self->file, "xf");
///     }
///     else {
///         lxw_xml_empty_tag(self->file, "xf", &attributes);
///     }

///     LXW_FREE_ATTRIBUTES();
/// }

std::string style_t::write_cell_xfs() const
{
  ///     struct xml_attribute_list attributes;
  ///     struct xml_attribute *attribute;
  ///     lxw_format *format;
  ///     uint32_t count = self->xf_count;
  ///     uint32_t i = 0;

  /* If the last format is "font_only" it is for the comment font and
   * shouldn't be counted. This is a workaround to get the last object
   * in the list since STAILQ_LAST() requires __containerof and isn't
   * ANSI compatible. */
  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
  ///         i++;
  ///         if (i == self->xf_count && format->font_only)
  ///             count--;
  ///     }

  std::string xml_data = xml_start_tag("cellXfs", {
                                                      {"count", "0" /* TODO count*/}
  });

  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
  ///         if (!format->font_only)
  ///             _write_xf(self, format);
  ///     }

  xml_data += xml_end_tag("cellXfs");

  return xml_data;
}

std::string style_t::write_cell_style(const std::string& name, uint8_t xf_id, uint8_t builtin_id) const
{
  return xml_empty_tag("cellStyle", {
                                        {"name",      name                      },
                                        {"xfId",      std::to_string(xf_id)     },
                                        {"builtinId", std::to_string(builtin_id)},
  });
}

std::string style_t::write_cell_styles() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  ///    if (self->has_hyperlink)
  ///        LXW_PUSH_ATTRIBUTES_STR("count", "2");
  ///    else
  attributes.emplace_back("count", "1");

  std::string xml_data = xml_start_tag("cellStyles", attributes);

  ///    if (self->has_hyperlink)
  ///        _write_cell_style(self, "Hyperlink", 1, 8);

  xml_data += write_cell_style("Normal", 0, 0);
  xml_data += xml_end_tag("cellStyles");

  return xml_data;
}

std::string style_t::write_dxfs() const
{
  ///     lxw_format *format;
  ///     uint32_t count = self->dxf_count;
  const std::vector<std::tuple<std::string, std::string>> attributes{
      {"count", "0" /* TODO count*/}
  };

  ///     if (count) {
  ///         lxw_xml_start_tag(self->file, "dxfs", &attributes);

  ///         STAILQ_FOREACH(format, self->dxf_formats, list_pointers) {
  ///             lxw_xml_start_tag(self->file, "dxf", NULL);

  ///             if (format->has_dxf_font)
  ///                 _write_font(self, format, LXW_TRUE, LXW_FALSE);

  ///             if (format->num_format_index)
  ///                 _write_num_fmt(self, format->num_format_index,
  ///                                format->num_format);

  ///             if (format->has_dxf_fill)
  ///                 _write_fill(self, format, LXW_TRUE);

  ///             if (format->has_dxf_border)
  ///                 _write_border(self, format, LXW_TRUE);

  ///             lxw_xml_end_tag(self->file, "dxf");
  ///         }

  ///         lxw_xml_end_tag(self->file, "dxfs");
  ///     }
  ///     else {
  return xml_empty_tag("dxfs", attributes);
  ///     }
}

std::string style_t::write_table_styles() const
{
  return xml_empty_tag("tableStyles", {
                                          {"count",             "0"                },
                                          {"defaultTableStyle", "TableStyleMedium9"},
                                          {"defaultPivotStyle", "PivotStyleLight16"},
  });
}

std::string style_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_style_sheet();
  xml_data += write_num_fmts();
  xml_data += write_fonts();
  xml_data += write_fills();
  xml_data += write_borders();
  xml_data += write_cell_style_xfs();
  xml_data += write_cell_xfs();
  xml_data += write_cell_styles();
  xml_data += write_dxfs();
  xml_data += write_table_styles();
  xml_data += xml_end_tag("styleSheet");

  return xml_data;
}

}
