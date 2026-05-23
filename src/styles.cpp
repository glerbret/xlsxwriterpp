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

namespace xwpp
{

/// STATIC void _write_font(lxw_styles *self, lxw_format *format, uint8_t is_dxf,
///                         uint8_t is_rich_string);

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

std::string styles_t::write_style_sheet() const
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

std::string styles_t::write_num_fmts() const
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

/// STATIC void
/// _write_font_size(lxw_styles *self, double font_size)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_DBL("val", font_size);

///     lxw_xml_empty_tag(self->file, "sz", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_color_theme(lxw_styles *self, uint8_t theme)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("theme", theme);

///     lxw_xml_empty_tag(self->file, "color", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_color_rgb(lxw_styles *self, int32_t rgb)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb_str[LXW_ATTR_32];

///     lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", rgb & LXW_COLOR_MASK);

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("rgb", rgb_str);

///     lxw_xml_empty_tag(self->file, "color", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_color_indexed(lxw_styles *self, uint8_t index)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("indexed", index);

///     lxw_xml_empty_tag(self->file, "color", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_name(lxw_styles *self, const char *font_name,
///                  uint8_t is_rich_string)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (*font_name)
///         LXW_PUSH_ATTRIBUTES_STR("val", font_name);
///     else
///         LXW_PUSH_ATTRIBUTES_STR("val", LXW_DEFAULT_FONT_NAME);

///     if (is_rich_string)
///         lxw_xml_empty_tag(self->file, "rFont", &attributes);
///     else
///         lxw_xml_empty_tag(self->file, "name", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_family(lxw_styles *self, uint8_t font_family)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("val", font_family);

///     lxw_xml_empty_tag(self->file, "family", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_charset(lxw_styles *self, uint8_t font_charset)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_INT("val", font_charset);

///     lxw_xml_empty_tag(self->file, "charset", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_scheme(lxw_styles *self, const char *font_scheme)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

///     if (*font_scheme)
///         LXW_PUSH_ATTRIBUTES_STR("val", font_scheme);
///     else
///         LXW_PUSH_ATTRIBUTES_STR("val", "minor");

///     lxw_xml_empty_tag(self->file, "scheme", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_underline(lxw_styles *self, uint8_t underline)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

/* Handle the underline variants. */
///     if (underline == LXW_UNDERLINE_DOUBLE)
///         LXW_PUSH_ATTRIBUTES_STR("val", "double");
///     else if (underline == LXW_UNDERLINE_SINGLE_ACCOUNTING)
///         LXW_PUSH_ATTRIBUTES_STR("val", "singleAccounting");
///     else if (underline == LXW_UNDERLINE_DOUBLE_ACCOUNTING)
///         LXW_PUSH_ATTRIBUTES_STR("val", "doubleAccounting");
/* Default to single underline. */

///     lxw_xml_empty_tag(self->file, "u", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_condense(lxw_styles *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("val", "0");

///     lxw_xml_empty_tag(self->file, "condense", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_extend(lxw_styles *self)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("val", "0");

///     lxw_xml_empty_tag(self->file, "extend", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font_vert_align(lxw_styles *self, const char *align)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("val", align);

///     lxw_xml_empty_tag(self->file, "vertAlign", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_font(lxw_styles *self, lxw_format *format, uint8_t is_dxf,
///             uint8_t is_rich_string)
/// {
///     if (is_rich_string)
///         lxw_xml_start_tag(self->file, "rPr", NULL);
///     else
///         lxw_xml_start_tag(self->file, "font", NULL);

///     if (format->font_condense)
///         _write_font_condense(self);

///     if (format->font_extend)
///         _write_font_extend(self);

///     if (format->bold)
///         lxw_xml_empty_tag(self->file, "b", NULL);

///     if (format->italic)
///         lxw_xml_empty_tag(self->file, "i", NULL);

///     if (format->font_strikeout)
///         lxw_xml_empty_tag(self->file, "strike", NULL);

///     if (format->font_outline)
///         lxw_xml_empty_tag(self->file, "outline", NULL);

///     if (format->font_shadow)
///         lxw_xml_empty_tag(self->file, "shadow", NULL);

///     if (format->underline)
///         _write_font_underline(self, format->underline);

///     if (format->font_script == LXW_FONT_SUPERSCRIPT)
///         _write_font_vert_align(self, "superscript");

///     if (format->font_script == LXW_FONT_SUBSCRIPT)
///         _write_font_vert_align(self, "subscript");

///     if (!is_dxf && format->font_size > 0.0)
///         _write_font_size(self, format->font_size);

///     if (format->theme)
///         _write_font_color_theme(self, format->theme);
///     else if (format->color_indexed)
///         _write_font_color_indexed(self, format->color_indexed);
///     else if (format->font_color != LXW_COLOR_UNSET)
///         _write_font_color_rgb(self, format->font_color);
///     else if (!is_dxf)
///         _write_font_color_theme(self, LXW_DEFAULT_FONT_THEME);

///     if (!is_dxf) {
///         _write_font_name(self, format->font_name, is_rich_string);
///         _write_font_family(self, format->font_family);

///         if (format->font_charset)
///             _write_font_charset(self, format->font_charset);

/* Only write the scheme element for the default font type if it
 * isn't a hyperlink. */
///         if ((!*format->font_name
///              || strcmp(LXW_DEFAULT_FONT_NAME, format->font_name) == 0)
///             && !format->hyperlink) {
///             _write_font_scheme(self, format->font_scheme);
///         }
///     }

///     if (format->hyperlink) {
///         self->has_hyperlink = LXW_TRUE;

///         if (self->hyperlink_font_id == 0)
///             self->hyperlink_font_id = format->font_index;
///     }

///     if (is_rich_string)
///         lxw_xml_end_tag(self->file, "rPr");
///     else
///         lxw_xml_end_tag(self->file, "font");
/// }

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

std::string styles_t::write_fonts() const
{
  ///    struct xml_attribute_list attributes;
  ///    struct xml_attribute *attribute;
  ///    lxw_format *format;
  const uint32_t count = 0;

  ///    count = self->font_count;
  ///    if (self->has_comments)
  ///        count++;

  std::string xml_data = xml_start_tag("fonts", {
                                                    {"count", std::to_string(count)}
  });

  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
  ///         if (format->has_font)
  ///             _write_font(self, format, LXW_FALSE, LXW_FALSE);
  ///     }

  ///     if (self->has_comments)
  ///         _write_comment_font(self);

  xml_data += xml_end_tag("fonts");

  return xml_data;
}

std::string styles_t::write_default_fill(const std::string& pattern) const
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

std::string styles_t::write_fills() const
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

/// STATIC void
/// _write_border_color(lxw_styles *self, lxw_color_t color)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;
///     char rgb_str[LXW_ATTR_32];

///     LXW_INIT_ATTRIBUTES();

///     if (color != LXW_COLOR_UNSET) {
///         lxw_snprintf(rgb_str, LXW_ATTR_32, "FF%06X", color & LXW_COLOR_MASK);
///         LXW_PUSH_ATTRIBUTES_STR("rgb", rgb_str);
///     }
///     else {
///         LXW_PUSH_ATTRIBUTES_STR("auto", "1");
///     }

///     lxw_xml_empty_tag(self->file, "color", &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_sub_border(lxw_styles *self, const char *type, uint8_t style,
///                   lxw_color_t color)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     char *border_styles[] = {
///         "none",
///         "thin",
///         "medium",
///         "dashed",
///         "dotted",
///         "thick",
///         "double",
///         "hair",
///         "mediumDashed",
///         "dashDot",
///         "mediumDashDot",
///         "dashDotDot",
///         "mediumDashDotDot",
///         "slantDashDot",
///     };

///     if (!style) {
///         lxw_xml_empty_tag(self->file, type, NULL);
///         return;
///     }

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("style", border_styles[style]);

///     lxw_xml_start_tag(self->file, type, &attributes);

///     _write_border_color(self, color);

///     lxw_xml_end_tag(self->file, type);

///     LXW_FREE_ATTRIBUTES();
/// }

/// STATIC void
/// _write_border(lxw_styles *self, lxw_format *format, uint8_t is_dxf)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();

/* Add attributes for diagonal borders. */
///     if (format->diag_type == LXW_DIAGONAL_BORDER_UP) {
///         LXW_PUSH_ATTRIBUTES_STR("diagonalUp", "1");
///     }
///     else if (format->diag_type == LXW_DIAGONAL_BORDER_DOWN) {
///         LXW_PUSH_ATTRIBUTES_STR("diagonalDown", "1");
///     }
///     else if (format->diag_type == LXW_DIAGONAL_BORDER_UP_DOWN) {
///         LXW_PUSH_ATTRIBUTES_STR("diagonalUp", "1");
///         LXW_PUSH_ATTRIBUTES_STR("diagonalDown", "1");
///     }

/* Ensure that a default diag border is set if the diag type is set. */
///     if (format->diag_type && !format->diag_border) {
///         format->diag_border = LXW_BORDER_THIN;
///     }

/* Write the start border tag. */
///     lxw_xml_start_tag(self->file, "border", &attributes);

/* Write the <border> sub elements. */
///     _write_sub_border(self, "left", format->left, format->left_color);
///     _write_sub_border(self, "right", format->right, format->right_color);
///     _write_sub_border(self, "top", format->top, format->top_color);
///     _write_sub_border(self, "bottom", format->bottom, format->bottom_color);

///     if (is_dxf) {
///         _write_sub_border(self, "vertical", 0, LXW_COLOR_UNSET);
///         _write_sub_border(self, "horizontal", 0, LXW_COLOR_UNSET);
///     }

/* Conditional DXF formats don't allow diagonal borders. */
///     if (!is_dxf)
///         _write_sub_border(self, "diagonal",
///                           format->diag_border, format->diag_color);

///     lxw_xml_end_tag(self->file, "border");

///     LXW_FREE_ATTRIBUTES();
/// }

std::string styles_t::write_borders() const
{
  ///     lxw_format *format;
  std::string xml_data = xml_start_tag("borders", {
                                                      {"count", "0" /*self->border_count*/}
  });
  ///     STAILQ_FOREACH(format, self->xf_formats, list_pointers) {
  ///         if (format->has_border)
  ///             _write_border(self, format, LXW_FALSE);
  ///     }

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

std::string styles_t::write_style_xf(bool has_hyperlink, uint16_t font_id) const
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

std::string styles_t::write_cell_style_xfs() const
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

std::string styles_t::write_cell_xfs() const
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

std::string styles_t::write_cell_style(const std::string& name, uint8_t xf_id, uint8_t builtin_id) const
{
  return xml_empty_tag("cellStyle", {
                                        {"name",      name                      },
                                        {"xfId",      std::to_string(xf_id)     },
                                        {"builtinId", std::to_string(builtin_id)},
  });
}

std::string styles_t::write_cell_styles() const
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

std::string styles_t::write_dxfs() const
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

std::string styles_t::write_table_styles() const
{
  return xml_empty_tag("tableStyles", {
                                          {"count",             "0"                },
                                          {"defaultTableStyle", "TableStyleMedium9"},
                                          {"defaultPivotStyle", "PivotStyleLight16"},
  });
}

std::string styles_t::assemble_xml_file() const
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
