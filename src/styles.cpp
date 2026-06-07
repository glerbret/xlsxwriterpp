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

#include <iostream>

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

std::string style_t::write_num_fmt(uint16_t num_fmt_id, const std::string& format_code) const
{
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

  std::vector<std::tuple<std::string, std::string>> attributes{
      {"numFmtId", std::to_string(num_fmt_id)}
  };
  ///     if (num_fmt_id < 50)
  ///         LXW_PUSH_ATTRIBUTES_STR("formatCode", format_codes[num_fmt_id]);
  ///     else if (num_fmt_id < 164)
  ///         LXW_PUSH_ATTRIBUTES_STR("formatCode", "General");
  ///     else
  attributes.emplace_back("formatCode", format_code);

  return xml_empty_tag("numFmt", attributes);
}

std::string style_t::write_num_fmts() const
{
  uint16_t last_format_index = 0;
  if(num_format_count_ == 0)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("numFmts", {
                                                      {"count", std::to_string(num_format_count_)}
  });

  // Write the numFmts elements.
  for(const auto format: xf_formats_)
  {
    // Ignore built-in number formats, i.e., < 0xA4.
    // TODO Add constantes for 0xA4
    if(format->num_format_index_ < 0xA4)
    {
      continue;
    }

    // Ignore duplicates which have an already used index.
    if(format->num_format_index_ <= last_format_index)
    {
      continue;
    }

    xml_data += write_num_fmt(format->num_format_index_, format->num_format_);
    last_format_index = format->num_format_index_;
  }

  xml_data += xml_end_tag("numFmts");

  return xml_data;
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
                                    {"rgb", std::format("FF{:06X}", static_cast<uint32_t>(rgb) & COLOR_MASK)}
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
      return "";
  }

  return "";
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

[[nodiscard]] std::string style_t::write_font(const format_t* format, bool is_dxf, bool is_rich_string)
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

  if(format->hyperlink_)
  {
    has_hyperlink_ = true;

    if(hyperlink_font_id_ == 0)
    {
      hyperlink_font_id_ = format->font_index_;
    }
  }

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

style_t::style_t(uint32_t font_count, uint32_t border_count, uint32_t num_format_count,
                 const std::vector<format_t*>& xf_formats)
  : font_count_{font_count}
  , num_format_count_{num_format_count}
  , border_count_{border_count}
  , xf_formats_{xf_formats}
{
}

std::string style_t::write_fonts()
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
      return "none";
  }

  return "none";
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

std::string style_t::write_hyperlink_alignment() const
{
  return xml_empty_tag("alignment", {
                                        {"vertical", "top"}
  });
}

std::string style_t::write_hyperlink_protection() const
{
  return xml_empty_tag("protection", {
                                         {"locked", "0"}
  });
}

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
    xml_data += write_hyperlink_alignment();
    xml_data += write_hyperlink_protection();
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

  if(has_hyperlink_)
  {
    attributes.emplace_back("count", "2");
  }
  else
  {
    attributes.emplace_back("count", "1");
  }

  std::string xml_data = xml_start_tag("cellStyleXfs", attributes);
  xml_data += write_style_xf(false, 0);

  if(has_hyperlink_)
  {
    xml_data += write_style_xf(has_hyperlink_, hyperlink_font_id_);
  }

  xml_data += xml_end_tag("cellStyleXfs");

  return xml_data;
}

bool style_t::apply_alignment(const format_t* format) const
{
  return format->text_h_align_ != format_alignments_t::NONE || format->text_v_align_ != format_alignments_t::NONE ||
         format->indent_ != 0 || format->rotation_ != 0 || format->text_wrap_ != 0 || format->shrink_ != 0 ||
         format->reading_order_ != 0;
}

bool style_t::has_alignment(const format_t* format) const
{
  return format->text_h_align_ != format_alignments_t::NONE ||
         !(format->text_v_align_ == format_alignments_t::NONE ||
           format->text_v_align_ == format_alignments_t::VERTICAL_BOTTOM) ||
         format->indent_ != 0 || format->rotation_ != 0 || format->text_wrap_ != 0 || format->shrink_ != 0 ||
         format->reading_order_ != 0;
}

std::string style_t::write_alignment(const format_t* format) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  int16_t rotation                 = format->rotation_;
  format_alignments_t text_h_align = format->text_h_align_;
  bool shrink                      = format->shrink_;
  bool just_distrib                = format->just_distrib_;

  /* Indent is only allowed for some alignment properties. */
  /* If it is defined for any other alignment or no alignment has been  */
  /* set then default to left alignment. */
  if(format->indent_ && text_h_align != format_alignments_t::HORIZONTAL_LEFT &&
     text_h_align != format_alignments_t::HORIZONTAL_RIGHT &&
     text_h_align != format_alignments_t::HORIZONTAL_DISTRIBUTED &&
     format->text_v_align_ != format_alignments_t::VERTICAL_TOP &&
     format->text_v_align_ != format_alignments_t::VERTICAL_BOTTOM &&
     format->text_v_align_ != format_alignments_t::VERTICAL_DISTRIBUTED)
  {
    text_h_align = format_alignments_t::HORIZONTAL_LEFT;
  }

  // Check for properties that are mutually exclusive.
  if(format->text_wrap_)
  {
    shrink = false;
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_FILL)
  {
    shrink = false;
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_JUSTIFY)
  {
    shrink = false;
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_DISTRIBUTED)
  {
    shrink = false;
  }

  if(text_h_align != format_alignments_t::HORIZONTAL_DISTRIBUTED)
  {
    just_distrib = 0;
  }

  if(format->indent_)
  {
    just_distrib = 0;
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_LEFT)
  {
    attributes.emplace_back("horizontal", "left");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_CENTER)
  {
    attributes.emplace_back("horizontal", "center");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_RIGHT)
  {
    attributes.emplace_back("horizontal", "right");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_FILL)
  {
    attributes.emplace_back("horizontal", "fill");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_JUSTIFY)
  {
    attributes.emplace_back("horizontal", "justify");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_CENTER_ACROSS)
  {
    attributes.emplace_back("horizontal", "centerContinuous");
  }

  if(text_h_align == format_alignments_t::HORIZONTAL_DISTRIBUTED)
  {
    attributes.emplace_back("horizontal", "distributed");
  }

  if(just_distrib)
  {
    attributes.emplace_back("justifyLastLine", "1");
  }

  if(format->text_v_align_ == format_alignments_t::VERTICAL_TOP)
  {
    attributes.emplace_back("vertical", "top");
  }

  if(format->text_v_align_ == format_alignments_t::VERTICAL_CENTER)
  {
    attributes.emplace_back("vertical", "center");
  }

  if(format->text_v_align_ == format_alignments_t::VERTICAL_JUSTIFY)
  {
    attributes.emplace_back("vertical", "justify");
  }

  if(format->text_v_align_ == format_alignments_t::VERTICAL_DISTRIBUTED)
  {
    attributes.emplace_back("vertical", "distributed");
  }

  // Map rotation to Excel values.
  if(rotation != 0)
  {
    if(rotation == 270)
    {
      rotation = 255;
    }
    else if(rotation < 0)
    {
      rotation = -rotation + 90;
    }

    attributes.emplace_back("textRotation", std::to_string(rotation));
  }

  if(format->indent_)
  {
    attributes.emplace_back("indent", std::to_string(format->indent_));
  }

  if(format->text_wrap_)
  {
    attributes.emplace_back("wrapText", "1");
  }

  if(shrink)
  {
    attributes.emplace_back("shrinkToFit", "1");
  }

  if(format->reading_order_ == 1)
  {
    attributes.emplace_back("readingOrder", "1");
  }

  if(format->reading_order_ == 2)
  {
    attributes.emplace_back("readingOrder", "2");
  }

  if(!attributes.empty())
  {
    return xml_empty_tag("alignment", attributes);
  }
  else
  {
    return "";
  }
}

std::string style_t::write_protection(const format_t* format) const
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  if(!format->locked_)
  {
    attributes.emplace_back("locked", "0");
  }

  if(format->hidden_)
  {
    attributes.emplace_back("hidden", "1");
  }

  return xml_empty_tag("protection", attributes);
}

std::string style_t::write_xf(const format_t* format) const
{
  const bool has_protection = (!format->locked_) | format->hidden_;

  std::vector<std::tuple<std::string, std::string>> attributes{
      {"numFmtId", std::to_string(format->num_format_index_)},
      {"fontId",   std::to_string(format->font_index_)      },
      {"fillId",   std::to_string(format->fill_index_)      },
      {"borderId", std::to_string(format->border_index_)    },
      {"xfId",     std::to_string(format->xf_id_)           },
  };

  if(format->quote_prefix_)
  {
    attributes.emplace_back("quotePrefix", "1");
  }

  if(format->num_format_index_ > 0)
  {
    attributes.emplace_back("applyNumberFormat", "1");
  }

  // Add applyFont attribute if XF format uses a font element.
  if(format->font_index_ > 0 && !format->hyperlink_)
  {
    attributes.emplace_back("applyFont", "1");
  }

  // Add applyFill attribute if XF format uses a fill element.
  if(format->fill_index_ > 0)
  {
    attributes.emplace_back("applyFill", "1");
  }

  // Add applyBorder attribute if XF format uses a border element.
  if(format->border_index_ > 0)
  {
    attributes.emplace_back("applyBorder", "1");
  }

  // We can also have applyAlignment without a sub-element.
  if(apply_alignment(format) || format->hyperlink_)
  {
    attributes.emplace_back("applyAlignment", "1");
  }

  if(has_protection || format->hyperlink_)
  {
    attributes.emplace_back("applyProtection", "1");
  }

  // Write XF with sub-elements if required.
  if(has_alignment(format) || has_protection)
  {
    std::string xml_data = xml_start_tag("xf", attributes);

    if(has_alignment(format))
    {
      xml_data += write_alignment(format);
    }

    if(has_protection)
    {
      xml_data += write_protection(format);
    }

    xml_data += xml_end_tag("xf");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("xf", attributes);
  }
}

std::string style_t::write_cell_xfs() const
{
  uint32_t count = xf_formats_.size();
  /* If the last format is "font_only" it is for the comment font and
   * shouldn't be counted. This is a workaround to get the last object
   * in the list since STAILQ_LAST() requires __containerof and isn't
   * ANSI compatible. */
  if(xf_formats_.back()->font_only_)
  {
    count--;
  }

  std::string xml_data = xml_start_tag("cellXfs", {
                                                      {"count", std::to_string(count)}
  });
  for(const auto format: xf_formats_)
  {
    if(!format->font_only_)
    {
      xml_data += write_xf(format);
    }
  }
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

  if(has_hyperlink_)
  {
    attributes.emplace_back("count", "2");
  }
  else
  {
    attributes.emplace_back("count", "1");
  }

  std::string xml_data = xml_start_tag("cellStyles", attributes);

  if(has_hyperlink_)
  {
    xml_data += write_cell_style("Hyperlink", 1, 8);
  }

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

// TODO Add const (remove to change hyperlink property)
std::string style_t::assemble_xml_file()
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
