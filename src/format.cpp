/*
 * format - A library for creating Excel XLSX format files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/format.h"

namespace xwpp
{

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

/// STATIC lxw_format * _get_format_key(lxw_format *self)
/// {
///   lxw_format *key = calloc(1, sizeof(lxw_format));
///   GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

///   memcpy(key, self, sizeof(lxw_format));

/* Set pointer members to NULL since they aren't part of the comparison. */
///   key->xf_format_indices = NULL;
///   key->dxf_format_indices = NULL;
///   key->num_xf_formats = NULL;
///   key->num_dxf_formats = NULL;
///   key->list_pointers.stqe_next = NULL;

///   return key;

/// mem_error:
///   return NULL;
/// }

/// lxw_border * lxw_format_get_border_key(lxw_format *self)
/// {
///   lxw_border *key = calloc(1, sizeof(lxw_border));
///   GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

///   key->bottom = self->bottom;
///   key->left = self->left;
///   key->right = self->right;
///   key->top = self->top;
///   key->diag_border = self->diag_border;
///   key->diag_type = self->diag_type;
///   key->bottom_color = self->bottom_color;
///   key->left_color = self->left_color;
///   key->right_color = self->right_color;
///   key->top_color = self->top_color;
///   key->diag_color = self->diag_color;

///   return key;

/// mem_error:
///   return NULL;
/// }

/// lxw_fill * lxw_format_get_fill_key(lxw_format *self)
/// {
///   lxw_fill *key = calloc(1, sizeof(lxw_fill));
///   GOTO_LABEL_ON_MEM_ERROR(key, mem_error);

///   key->fg_color = self->fg_color;
///   key->bg_color = self->bg_color;
///   key->pattern = self->pattern;

///   return key;

/// mem_error:
///   return NULL;
/// }

/// int32_t lxw_format_get_dxf_index(lxw_format *self)
/// {
///   lxw_format *format_key;
///   lxw_format *existing_format;
///   lxw_hash_element *hash_element;
///   lxw_hash_table *formats_hash_table = self->dxf_format_indices;
///   int32_t index;

/* Note: The formats_hash_table/dxf_format_indices contains the unique and
 * more importantly the *used* formats in the workbook.
 */

/* Format already has an index number so return it. */
///   if (self->dxf_index != LXW_PROPERTY_UNSET) {
///     return self->dxf_index;
///   }

/* Otherwise, the format doesn't have an index number so we assign one.
 * First generate a unique key to identify the format in the hash table.
 */
///   format_key = _get_format_key(self);

/* Return the default format index if the key generation failed. */
///   if (!format_key)
///     return 0;

/* Look up the format in the hash table. */
///   hash_element = lxw_hash_key_exists(formats_hash_table, format_key, sizeof(lxw_format));

///   if (hash_element) {
/* Format matches existing format with an index. */
///     free(format_key);
///     existing_format = hash_element->value;
///     return existing_format->dxf_index;
///   }
///   else {
/* New format requiring an index. */
///     index = formats_hash_table->unique_count;
///     self->dxf_index = index;
///     lxw_insert_hash_element(formats_hash_table, format_key, self, sizeof(lxw_format));
///     return index;
///   }
/// }

/// void format_set_font_name(lxw_format *self, const char *font_name)
/// {
///   LXW_FORMAT_FIELD_COPY(self->font_name, font_name);
/// }

/// void format_set_font_size(lxw_format *self, double size)
/// {
///   if (size >= LXW_MIN_FONT_SIZE && size <= LXW_MAX_FONT_SIZE)
///     self->font_size = size;
/// }

void format_t::set_font_color(color_t color)
{
  font_color_ = color;
}

void format_t::set_bold()
{
  bold_ = true;
}

void format_t::set_italic()
{
  italic_ = true;
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

/// void format_set_font_strikeout(lxw_format *self)
/// {
///   self->font_strikeout = LXW_TRUE;
/// }

/// void format_set_font_script(lxw_format *self, uint8_t style)
/// {
///   if (style >= LXW_FONT_SUPERSCRIPT && style <= LXW_FONT_SUBSCRIPT)
///     self->font_script = style;
/// }

/// void format_set_font_outline(lxw_format *self)
/// {
///   self->font_outline = LXW_TRUE;
/// }

/// void format_set_font_shadow(lxw_format *self)
/// {
///   self->font_shadow = LXW_TRUE;
/// }

void format_t::set_num_format(const std::string& num_format)
{
  num_format_ = num_format;
}

/// void format_set_unlocked(lxw_format *self)
/// {
///   self->locked = LXW_FALSE;
/// }

/// void format_set_hidden(lxw_format *self)
/// {
///   self->hidden = LXW_TRUE;
/// }

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

void format_t::set_text_wrap()
{
  text_wrap_ = true;
}

/// void format_set_rotation(lxw_format *self, int16_t angle)
/// {
/* Convert user angle to Excel angle. */
///   if (angle == 270) {
///     self->rotation = 255;
///   }
///   else if (angle >= -90 && angle <= 90) {
///     if (angle < 0)
///       angle = -angle + 90;

///     self->rotation = angle;
///   }
///   else {
///     LXW_WARN("Rotation rotation outside range: -90 <= angle <= 90.");
///     self->rotation = 0;
///   }
/// }

/// void format_set_indent(lxw_format *self, uint8_t value)
/// {
///   self->indent = value;
/// }

/// void format_set_shrink(lxw_format *self)
/// {
///   self->shrink = LXW_TRUE;
/// }

/// void format_set_text_justlast(lxw_format *self)
/// {
///   self->text_justlast = LXW_TRUE;
/// }

void format_t::set_pattern(format_patterns_t pattern)
{
  ///   if (value > LXW_PATTERN_GRAY_0625) {
  ///     LXW_WARN_FORMAT1("format_set_pattern(): invalid pattern value: %d", value);
  ///     return;
  ///   }

  pattern_ = pattern;
}

void format_t::set_bg_color(color_t color)
{
  bg_color_ = color;
}

void format_t::set_fg_color(color_t color)
{
  fg_color_ = color;
}

void format_t::set_border(format_borders_t style)
{
  style = check_border(style);

  bottom_ = style;
  top_    = style;
  left_   = style;
  right_  = style;
}

void format_t::set_border_color(color_t color)
{
  bottom_color_ = color;
  top_color_    = color;
  left_color_   = color;
  right_color_  = color;
}

void format_t::set_bottom(format_borders_t style)
{
  bottom_ = check_border(style);
}

void format_t::set_bottom_color(color_t color)
{
  bottom_color_ = color;
}

void format_t::set_left(format_borders_t style)
{
  left_ = check_border(style);
}

void format_t::set_left_color(color_t color)
{
  left_color_ = color;
}

void format_t::set_right(format_borders_t style)
{
  right_ = check_border(style);
}

void format_t::set_right_color(color_t color)
{
  right_color_ = color;
}

void format_t::set_top(format_borders_t style)
{
  top_ = check_border(style);
}

void format_t::set_top_color(color_t color)
{
  top_color_ = color;
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

void format_t::set_num_format_index(uint8_t value)
{
  num_format_index_ = value;
}

/// void format_set_valign(lxw_format *self, uint8_t value)
/// {
///   if (value > LXW_ALIGN_VERTICAL_DISTRIBUTED) {
///     LXW_WARN_FORMAT1
///       ("format_set_valign(): invalid vertical alignment value: %d",
///        value);
///     return;
///   }

///   self->text_v_align = value;
/// }

/// void format_set_reading_order(lxw_format *self, uint8_t value)
/// {
///   self->reading_order = value;
/// }

/// void format_set_font_family(lxw_format *self, uint8_t value)
/// {
///   self->font_family = value;
/// }

/// void format_set_font_charset(lxw_format *self, uint8_t value)
/// {
///   self->font_charset = value;
/// }

/// void format_set_font_scheme(lxw_format *self, const char *font_scheme)
/// {
///   LXW_FORMAT_FIELD_COPY(self->font_scheme, font_scheme);
/// }

/// void format_set_font_condense(lxw_format *self)
/// {
///   self->font_condense = LXW_TRUE;
/// }

/// void format_set_font_extend(lxw_format *self)
/// {
///   self->font_extend = LXW_TRUE;
/// }

/// void format_set_theme(lxw_format *self, uint8_t value)
/// {
///   self->theme = value;
/// }

/// void format_set_color_indexed(lxw_format *self, uint8_t value)
/// {
///   self->color_indexed = value;
/// }

/// void format_set_font_only(lxw_format *self)
/// {
///   self->font_only = LXW_TRUE;
/// }

void format_t::set_hyperlink()
{
  hyperlink_ = true;
  xf_id_     = 1;
  underline_ = format_underlines_t::SINGLE;
  theme_     = 10;
}

/// void format_set_quote_prefix(lxw_format *self)
/// {
///   self->quote_prefix = LXW_TRUE;
/// }

const std::string format_t::DEFAULT_FONT_NAME = "Calibri";

}
