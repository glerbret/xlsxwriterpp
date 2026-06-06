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

/// STATIC uint8_t _check_border(uint8_t border)
/// {
///   if (border >= LXW_BORDER_THIN && border <= LXW_BORDER_SLANT_DASH_DOT)
///     return border;
///   else
///     return LXW_BORDER_NONE;
/// }

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

/// void format_set_font_color(lxw_format *self, lxw_color_t color)
/// {
///   self->font_color = color;
/// }

/// void format_set_bold(lxw_format *self)
/// {
///   self->bold = LXW_TRUE;
/// }

/// void format_set_italic(lxw_format *self)
/// {
///   self->italic = LXW_TRUE;
/// }

/// void format_set_underline(lxw_format *self, uint8_t style)
/// {
///   if (style >= LXW_UNDERLINE_SINGLE && style <= LXW_UNDERLINE_DOUBLE_ACCOUNTING)
///     self->underline = style;
/// }

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

/// void format_set_num_format(lxw_format *self, const char *num_format)
/// {
///   LXW_FORMAT_FIELD_COPY(self->num_format, num_format);
/// }

/// void format_set_unlocked(lxw_format *self)
/// {
///   self->locked = LXW_FALSE;
/// }

/// void format_set_hidden(lxw_format *self)
/// {
///   self->hidden = LXW_TRUE;
/// }

/// void format_set_align(lxw_format *self, uint8_t value)
/// {
///   if (value >= LXW_ALIGN_LEFT && value <= LXW_ALIGN_DISTRIBUTED) {
///     self->text_h_align = value;
///   }

///   if (value >= LXW_ALIGN_VERTICAL_TOP && value <= LXW_ALIGN_VERTICAL_DISTRIBUTED) {
///     self->text_v_align = value;
///   }
/// }

/// void format_set_text_wrap(lxw_format *self)
/// {
///   self->text_wrap = LXW_TRUE;
/// }

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

/// void format_set_pattern(lxw_format *self, uint8_t value)
/// {
///   if (value > LXW_PATTERN_GRAY_0625) {
///     LXW_WARN_FORMAT1("format_set_pattern(): invalid pattern value: %d", value);
///     return;
///   }

///   self->pattern = value;
/// }

/// void format_set_bg_color(lxw_format *self, lxw_color_t color)
/// {
///   self->bg_color = color;
/// }

/// void format_set_fg_color(lxw_format *self, lxw_color_t color)
/// {
///   self->fg_color = color;
/// }

/// void format_set_border(lxw_format *self, uint8_t style)
/// {
///   style = _check_border(style);
///   self->bottom = style;
///   self->top = style;
///   self->left = style;
///   self->right = style;
/// }

/// void format_set_border_color(lxw_format *self, lxw_color_t color)
/// {
///   self->bottom_color = color;
///   self->top_color = color;
///   self->left_color = color;
///   self->right_color = color;
/// }

/// void format_set_bottom(lxw_format *self, uint8_t style)
/// {
///   self->bottom = _check_border(style);
/// }

/// void format_set_bottom_color(lxw_format *self, lxw_color_t color)
/// {
///   self->bottom_color = color;
/// }

/// void format_set_left(lxw_format *self, uint8_t style)
/// {
///   self->left = _check_border(style);
/// }

/// void format_set_left_color(lxw_format *self, lxw_color_t color)
/// {
///   self->left_color = color;
/// }

/// void format_set_right(lxw_format *self, uint8_t style)
/// {
///   self->right = _check_border(style);
/// }

/// void format_set_right_color(lxw_format *self, lxw_color_t color)
/// {
///   self->right_color = color;
/// }

/// void format_set_top(lxw_format *self, uint8_t style)
/// {
///   self->top = _check_border(style);
/// }

/// void format_set_top_color(lxw_format *self, lxw_color_t color)
/// {
///   self->top_color = color;
/// }

/// void format_set_diag_type(lxw_format *self, uint8_t type)
/// {
///   if (type >= LXW_DIAGONAL_BORDER_UP && type <= LXW_DIAGONAL_BORDER_UP_DOWN)
///     self->diag_type = type;
/// }

/// void format_set_diag_color(lxw_format *self, lxw_color_t color)
/// {
///   self->diag_color = color;
/// }

/// void format_set_diag_border(lxw_format *self, uint8_t style)
/// {
///   if (style > LXW_BORDER_SLANT_DASH_DOT) {
///     LXW_WARN_FORMAT1("format_set_diag_border(): invalid border style: %d", style);
///     return;
///   }

///   self->diag_border = style;
/// }

/// void format_set_num_format_index(lxw_format *self, uint8_t value)
/// {
///   self->num_format_index = value;
/// }

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

/// void format_set_hyperlink(lxw_format *self)
/// {
///   self->hyperlink = LXW_TRUE;
///   self->xf_id = 1;
///   self->underline = LXW_UNDERLINE_SINGLE;
///   self->theme = 10;
/// }

/// void format_set_quote_prefix(lxw_format *self)
/// {
///   self->quote_prefix = LXW_TRUE;
/// }

const std::string format_t::DEFAULT_FONT_NAME = "Calibri";

}
