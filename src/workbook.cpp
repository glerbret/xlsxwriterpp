/*
 * workbook - A library for creating Excel XLSX workbook files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 */

#include "xwpp/workbook.h"

#include "xwpp/exception.h"
#include "xwpp/packager.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <chrono>
#include <format>
#include <string>

#include <iostream>

using namespace std::literals::chrono_literals;

// TODO Add API to configure default format, in particular font

namespace xwpp
{

/// STATIC int _worksheet_name_cmp(lxw_worksheet_name *name1,
///                                lxw_worksheet_name *name2);
/// STATIC int _chartsheet_name_cmp(lxw_chartsheet_name *name1,
///                                 lxw_chartsheet_name *name2);
/// STATIC int _image_md5_cmp(lxw_image_md5 *tuple1, lxw_image_md5 *tuple2);

/// STATIC int
/// _worksheet_name_cmp(lxw_worksheet_name *name1, lxw_worksheet_name *name2)
/// {
///     return lxw_strcasecmp(name1->name, name2->name);
/// }

/// STATIC int
/// _chartsheet_name_cmp(lxw_chartsheet_name *name1, lxw_chartsheet_name *name2)
/// {
///     return lxw_strcasecmp(name1->name, name2->name);
/// }

/// STATIC int
/// _image_md5_cmp(lxw_image_md5 *tuple1, lxw_image_md5 *tuple2)
/// {
///     return strcmp(tuple1->md5, tuple2->md5);
/// }

/// STATIC void  _free_custom_doc_property(lxw_custom_property *custom_property)
/// {
///     if (custom_property) {
///         free(custom_property->name);
///         if (custom_property->type == LXW_CUSTOM_STRING)
///             free(custom_property->u.string);
///     }

///     free(custom_property);
/// }

/// void lxw_workbook_set_default_xf_indices(lxw_workbook *self)
/// {
///     lxw_format *format;
///     int32_t index = 0;

///     STAILQ_FOREACH(format, self->formats, list_pointers) {

/* Skip the hyperlink format. */
///         if (index != 1)
///             lxw_format_get_xf_index(format);

///         index++;
///     }
/// }

// ICI
void workbook_t::prepare_fonts()
{
  std::vector<format_t*> fonts;
  ///     lxw_hash_table *fonts = lxw_hash_new(128, 1, 1);
  ///     lxw_hash_element *hash_element;
  ///     lxw_hash_element *used_format_element;
  ///     uint16_t index = 0;

  // TODO Use unordered_set to optimise this search
  for(const auto format: used_xf_formats_)
  {
    for(const auto font: fonts)
    {
      if(format->font_name_ == font->font_name_ && format->font_size_ == font->font_size_ &&
         format->bold_ == font->bold_ && format->italic_ == font->italic_ && format->underline_ == font->underline_ &&
         format->theme_ == font->theme_ && format->font_color_ == font->font_color_ &&
         format->font_strikeout_ == font->font_strikeout_ && format->font_outline_ == font->font_outline_ &&
         format->font_shadow_ == font->font_shadow_ && format->font_script_ == font->font_script_ &&
         format->font_family_ == font->font_family_ && format->font_charset_ == font->font_charset_ &&
         format->font_condense_ == font->font_condense_ && format->font_extend_ == font->font_extend_)
      {
        // Font has already been used.
        format->font_index_ = font->font_index_;
        format->has_font_   = false;
      }
    }
    if(format->font_index_ == format_t::PROPERTY_UNSET)
    {
      format->font_index_ = static_cast<int32_t>(fonts.size());
      format->has_font_   = true;
      fonts.push_back(format);
    }
  }

  /* For DXF formats we only need to check if the properties have changed. */
  ///     LXW_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
  ///         lxw_format *format = (lxw_format *) used_format_element->value;

  /* The only font properties that can change for a DXF format are:
   * color, bold, italic, underline and strikethrough. */
  ///         if (format->font_color || format->bold || format->italic
  ///             || format->underline || format->font_strikeout) {
  ///             format->has_dxf_font = LXW_TRUE;
  ///         }
  ///     }

  font_count_ = static_cast<uint32_t>(fonts.size());
}

/// STATIC void _prepare_borders(lxw_workbook *self)
/// {
///     lxw_hash_table *borders = lxw_hash_new(128, 1, 1);
///     lxw_hash_element *hash_element;
///     lxw_hash_element *used_format_element;
///     uint16_t index = 0;

///     LXW_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;
///         lxw_border *key = lxw_format_get_border_key(format);

///         if (key) {
/* Look up the format in the hash table. */
///             hash_element =
///                 lxw_hash_key_exists(borders, key, sizeof(lxw_border));

///             if (hash_element) {
/* Border has already been used. */
///                 format->border_index = *(uint16_t *) hash_element->value;
///                 format->has_border = LXW_FALSE;
///                 free(key);
///             }
///             else {
/* This is a new border. */
///                 uint16_t *border_index = calloc(1, sizeof(uint16_t));
///                 *border_index = index;
///                 format->border_index = index;
///                 format->has_border = 1;
///                 lxw_insert_hash_element(borders, key, border_index,
///                                         sizeof(lxw_border));
///                 index++;
///             }
///         }
///     }

/* For DXF formats we only need to check if the properties have changed. */
///     LXW_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;

///         if (format->left || format->right || format->top || format->bottom) {
///             format->has_dxf_border = LXW_TRUE;
///         }
///     }

///     lxw_hash_free(borders);

///     self->border_count = index;
/// }

/// STATIC void _prepare_fills(lxw_workbook *self)
/// {
///     lxw_hash_table *fills = lxw_hash_new(128, 1, 1);
///     lxw_hash_element *hash_element;
///     lxw_hash_element *used_format_element;
///     uint16_t index = 2;
///     lxw_fill *default_fill_1 = NULL;
///     lxw_fill *default_fill_2 = NULL;
///     uint16_t *fill_index1 = NULL;
///     uint16_t *fill_index2 = NULL;

///     default_fill_1 = calloc(1, sizeof(lxw_fill));
///     GOTO_LABEL_ON_MEM_ERROR(default_fill_1, mem_error);

///     default_fill_2 = calloc(1, sizeof(lxw_fill));
///     GOTO_LABEL_ON_MEM_ERROR(default_fill_2, mem_error);

///     fill_index1 = calloc(1, sizeof(uint16_t));
///     GOTO_LABEL_ON_MEM_ERROR(fill_index1, mem_error);

///     fill_index2 = calloc(1, sizeof(uint16_t));
///     GOTO_LABEL_ON_MEM_ERROR(fill_index2, mem_error);

/* Add the default fills. */
///     default_fill_1->pattern = LXW_PATTERN_NONE;
///     default_fill_1->fg_color = LXW_COLOR_UNSET;
///     default_fill_1->bg_color = LXW_COLOR_UNSET;
///     *fill_index1 = 0;
///     lxw_insert_hash_element(fills, default_fill_1, fill_index1,
///                             sizeof(lxw_fill));

///     default_fill_2->pattern = LXW_PATTERN_GRAY_125;
///     default_fill_2->fg_color = LXW_COLOR_UNSET;
///     default_fill_2->bg_color = LXW_COLOR_UNSET;
///     *fill_index2 = 1;
///     lxw_insert_hash_element(fills, default_fill_2, fill_index2,
///                             sizeof(lxw_fill));

/* For DXF formats we only need to check if the properties have changed. */
///     LXW_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;

///         if (format->pattern || format->bg_color || format->fg_color) {
///             format->has_dxf_fill = LXW_TRUE;
///             format->dxf_bg_color = format->bg_color;
///             format->dxf_fg_color = format->fg_color;
///         }
///     }

///     LXW_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;
///         lxw_fill *key = lxw_format_get_fill_key(format);

/* The following logical statements jointly take care of special */
/* cases in relation to cell colors and patterns:                */
/* 1. For a solid fill (pattern == 1) Excel reverses the role of */
/*    foreground and background colors, and                      */
/* 2. If the user specifies a foreground or background color     */
/*    without a pattern they probably wanted a solid fill, so    */
/*    we fill in the defaults.                                   */
///         if (format->pattern == LXW_PATTERN_SOLID
///             && format->bg_color != LXW_COLOR_UNSET
///             && format->fg_color != LXW_COLOR_UNSET) {
///             lxw_color_t tmp = format->fg_color;
///             format->fg_color = format->bg_color;
///             format->bg_color = tmp;
///         }

///         if (format->pattern <= LXW_PATTERN_SOLID
///             && format->bg_color != LXW_COLOR_UNSET
///             && format->fg_color == LXW_COLOR_UNSET) {
///             format->fg_color = format->bg_color;
///             format->bg_color = LXW_COLOR_UNSET;
///             format->pattern = LXW_PATTERN_SOLID;
///         }

///         if (format->pattern <= LXW_PATTERN_SOLID
///             && format->bg_color == LXW_COLOR_UNSET
///             && format->fg_color != LXW_COLOR_UNSET) {
///             format->pattern = LXW_PATTERN_SOLID;
///         }

///         if (key) {
/* Look up the format in the hash table. */
///             hash_element = lxw_hash_key_exists(fills, key, sizeof(lxw_fill));

///             if (hash_element) {
/* Fill has already been used. */
///                 format->fill_index = *(uint16_t *) hash_element->value;
///                 format->has_fill = LXW_FALSE;
///                 free(key);
///             }
///             else {
/* This is a new fill. */
///                 uint16_t *fill_index = calloc(1, sizeof(uint16_t));
///                 *fill_index = index;
///                 format->fill_index = index;
///                 format->has_fill = 1;
///                 lxw_insert_hash_element(fills, key, fill_index,
///                                         sizeof(lxw_fill));
///                 index++;
///             }
///         }
///     }

///     lxw_hash_free(fills);

///     self->fill_count = index;

///     return;

/// mem_error:
///     free(fill_index2);
///     free(fill_index1);
///     free(default_fill_2);
///     free(default_fill_1);
///     lxw_hash_free(fills);
/// }

/// STATIC void
/// _prepare_num_formats(lxw_workbook *self)
/// {
///     lxw_hash_table *num_formats = lxw_hash_new(128, 0, 1);
///     lxw_hash_element *hash_element;
///     lxw_hash_element *used_format_element;
///     uint16_t index = 0xA4;
///     uint16_t num_format_count = 0;
///     uint16_t *num_format_index;

///     LXW_FOREACH_ORDERED(used_format_element, self->used_xf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;

/* Format already has a number format index. */
///         if (format->num_format_index)
///             continue;

/* Check if there is a user defined number format string. */
///         if (*format->num_format) {
///             char num_format[LXW_FORMAT_FIELD_LEN] = { 0 };
///             lxw_snprintf(num_format, LXW_FORMAT_FIELD_LEN, "%s",
///                          format->num_format);

/* Look up the num_format in the hash table. */
///             hash_element = lxw_hash_key_exists(num_formats, num_format,
///                                                LXW_FORMAT_FIELD_LEN);

///             if (hash_element) {
/* Num_Format has already been used. */
///                 format->num_format_index = *(uint16_t *) hash_element->value;
///             }
///             else {
/* This is a new num_format. */
///                 num_format_index = calloc(1, sizeof(uint16_t));
///                 *num_format_index = index;
///                 format->num_format_index = index;
///                 lxw_insert_hash_element(num_formats, format->num_format,
///                                         num_format_index,
///                                         LXW_FORMAT_FIELD_LEN);
///                 index++;
///                 num_format_count++;
///             }
///         }
///     }

///     LXW_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
///         lxw_format *format = (lxw_format *) used_format_element->value;

/* Format already has a number format index. */
///         if (format->num_format_index)
///             continue;

/* Check if there is a user defined number format string. */
///         if (*format->num_format) {
///             char num_format[LXW_FORMAT_FIELD_LEN] = { 0 };
///             lxw_snprintf(num_format, LXW_FORMAT_FIELD_LEN, "%s",
///                          format->num_format);

/* Look up the num_format in the hash table. */
///             hash_element = lxw_hash_key_exists(num_formats, num_format,
///                                                LXW_FORMAT_FIELD_LEN);

///             if (hash_element) {
/* Num_Format has already been used. */
///                 format->num_format_index = *(uint16_t *) hash_element->value;
///             }
///             else {
/* This is a new num_format. */
///                 num_format_index = calloc(1, sizeof(uint16_t));
///                 *num_format_index = index;
///                 format->num_format_index = index;
///                 lxw_insert_hash_element(num_formats, format->num_format,
///                                         num_format_index,
///                                         LXW_FORMAT_FIELD_LEN);
///                 index++;
/* Don't update num_format_count for DXF formats. */
///             }
///         }
///     }

///     lxw_hash_free(num_formats);

///     self->num_format_count = num_format_count;
/// }

void workbook_t::prepare_workbook()
{
  // Set the font index for the format objects.
  prepare_fonts();

  // Set the number format index for the format objects.
  ///     _prepare_num_formats(self);

  // Set the border index for the format objects.
  ///     _prepare_borders(self);

  // Set the fill index for the format objects.
  ///     _prepare_fills(self);
}

/// static int _compare_defined_names(lxw_defined_name *a, lxw_defined_name *b)
/// {
///     int res = strcmp(a->normalised_name, b->normalised_name);

/* Primary comparison based on defined name. */
///     if (res)
///         return res;

/* Secondary comparison based on worksheet name. */
///     res = strcmp(a->normalised_sheetname, b->normalised_sheetname);

///     return res;
/// }

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
///  STATIC lxw_error _store_defined_name(lxw_workbook *self, const char *name,
///                     const char *app_name, const char *formula, int16_t index,
///                     uint8_t hidden)
/// {
///     lxw_sheet *sheet;
///     lxw_worksheet *worksheet;
///     lxw_defined_name *defined_name;
///     lxw_defined_name *list_defined_name;
///     char name_copy[LXW_DEFINED_NAME_LENGTH];
///     char *tmp_str;
///     char *worksheet_name;

/* Do some checks on the input data */
///     if (!name || !formula)
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;

///     if (lxw_str_is_empty(name) || lxw_str_is_empty(formula))
///         return LXW_ERROR_PARAMETER_IS_EMPTY;

///     if (lxw_utf8_strlen(name) > LXW_DEFINED_NAME_LENGTH ||
///         lxw_utf8_strlen(formula) > LXW_DEFINED_NAME_LENGTH) {
///         return LXW_ERROR_128_STRING_LENGTH_EXCEEDED;
///     }

/* Allocate a new defined_name to be added to the linked list of names. */
///     defined_name = calloc(1, sizeof(struct lxw_defined_name));
///     RETURN_ON_MEM_ERROR(defined_name, LXW_ERROR_MEMORY_MALLOC_FAILED);

/* Copy the user input string. */
///     lxw_strcpy(name_copy, name);

/* Set the worksheet index or -1 for a global defined name. */
///     defined_name->index = index;
///     defined_name->hidden = hidden;

/* Check for local defined names like like "Sheet1!name". */
///     tmp_str = strchr(name_copy, '!');

///     if (tmp_str == NULL) {
/* The name is global. We just store the defined name string. */
///         lxw_strcpy(defined_name->name, name_copy);
///     }
///     else {
/* The name is worksheet local. We need to extract the sheet name
 * and map it to a sheet index. */

/* Split the into the worksheet name and defined name. */
///         *tmp_str = '\0';
///         tmp_str++;
///         worksheet_name = name_copy;

///         if (lxw_str_is_empty(tmp_str) || lxw_str_is_empty(worksheet_name))
///             goto mem_error;

/* Remove any worksheet quoting. */
///         if (worksheet_name[0] == '\'')
///             worksheet_name++;
///         if (strlen(worksheet_name) > 0
///             && worksheet_name[strlen(worksheet_name) - 1] == '\'') {
///             worksheet_name[strlen(worksheet_name) - 1] = '\0';
///         }

/* Search for worksheet name to get the equivalent worksheet index. */
///         STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
///             if (sheet->is_chartsheet)
///                 continue;
///             else
///                 worksheet = sheet->u.worksheet;

///             if (strcmp(worksheet_name, worksheet->name) == 0) {
///                 defined_name->index = worksheet->index;
///                 lxw_strcpy(defined_name->normalised_sheetname,
///                            worksheet_name);
///             }
///         }

/* If we didn't find the worksheet name we exit. */
///         if (defined_name->index == -1)
///             goto mem_error;

///         lxw_strcpy(defined_name->name, tmp_str);
///     }

/* Print titles and repeat title pass in the name used for App.xml. */
///     if (app_name) {
///         lxw_strcpy(defined_name->app_name, app_name);
///         lxw_strcpy(defined_name->normalised_sheetname, app_name);
///     }
///     else {
///         lxw_strcpy(defined_name->app_name, name);
///     }

/* We need to normalize the defined names for sorting. This involves
 * removing any _xlnm namespace  and converting it to lowercase. */
///     tmp_str = strstr(name_copy, "_xlnm.");
///     if (tmp_str)
///         lxw_strcpy(defined_name->normalised_name, defined_name->name + 6);
///     else
///         lxw_strcpy(defined_name->normalised_name, defined_name->name);

///     lxw_str_tolower(defined_name->normalised_name);
///     lxw_str_tolower(defined_name->normalised_sheetname);

/* Strip leading "=" from the formula. */
///     if (formula[0] == '=')
///         lxw_strcpy(defined_name->formula, formula + 1);
///     else
///         lxw_strcpy(defined_name->formula, formula);

/* We add the defined name to the list in sorted order. */
///     list_defined_name = TAILQ_FIRST(self->defined_names);

///     if (list_defined_name == NULL ||
///         _compare_defined_names(defined_name, list_defined_name) < 1) {
/* List is empty or defined name goes to the head. */
///         TAILQ_INSERT_HEAD(self->defined_names, defined_name, list_pointers);
///         return LXW_NO_ERROR;
///     }

///     TAILQ_FOREACH(list_defined_name, self->defined_names, list_pointers) {
///         int res = _compare_defined_names(defined_name, list_defined_name);

/* The entry already exists. We exit and don't overwrite. */
///         if (res == 0)
///             goto mem_error;

/* New defined name is inserted in sorted order before other entries. */
///         if (res < 0) {
///             TAILQ_INSERT_BEFORE(list_defined_name, defined_name,
///                                 list_pointers);
///             return LXW_NO_ERROR;
///         }
/// }

/* If the entry wasn't less than any of the entries in the list we add it
 * to the end. */
///     TAILQ_INSERT_TAIL(self->defined_names, defined_name, list_pointers);
///     return LXW_NO_ERROR;

/// mem_error:
///     free(defined_name);
///     return LXW_ERROR_MEMORY_MALLOC_FAILED;
/// }

/*
 * Populate the data cache of a chart data series by reading the data from the
 * relevant worksheet and adding it to the cached in the range object as a
 * list of points.
 *
 * Note, the data cache isn't strictly required by Excel but it helps if the
 * chart is embedded in another application such as PowerPoint and it also
 * helps with comparison testing.
 */
/// STATIC void _populate_range_data_cache(lxw_workbook *self, lxw_series_range *range)
/// {
///     lxw_worksheet *worksheet;
///     row_num_t row_num;
///     col_num_t col_num;
///     row_t *row_obj;
///     cell_t *cell_obj;
///     struct lxw_series_data_point *data_point;
///     uint16_t num_data_points = 0;

/* If ignore_cache is set then don't try to populate the cache. This flag
 * may be set manually, for testing, or due to a case where the cache
 * can't be calculated.
 */
///     if (range->ignore_cache)
///         return;

/* Currently we only handle 2D ranges so ensure either the rows or cols
 * are the same.
 */
///     if (range->first_row != range->last_row
///         && range->first_col != range->last_col) {
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }

/* Check that the sheetname exists. */
///     worksheet = workbook_get_worksheet_by_name(self, range->sheetname);
///     if (!worksheet) {
///         LXW_WARN_FORMAT2("workbook_add_chart(): worksheet name '%s' "
///                          "in chart formula '%s' doesn't exist.",
///                          range->sheetname, range->formula);
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }

/* We can't read the data when worksheet optimization is on. */
///     if (worksheet->optimize) {
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }

/* Iterate through the worksheet data and populate the range cache. */
///     for (row_num = range->first_row; row_num <= range->last_row; row_num++) {
///         row_obj = lxw_worksheet_find_row(worksheet, row_num);

///         for (col_num = range->first_col; col_num <= range->last_col;
///              col_num++) {

///             data_point = calloc(1, sizeof(struct lxw_series_data_point));
///             if (!data_point) {
///                 range->ignore_cache = LXW_TRUE;
///                 return;
///             }

///             cell_obj = lxw_worksheet_find_cell_in_row(row_obj, col_num);

///             if (cell_obj) {
///                 if (cell_obj->type == NUMBER_CELL) {
///                     data_point->number = cell_obj->u.number;
///                 }

///                 if (cell_obj->type == STRING_CELL) {
///                     data_point->string = lxw_strdup(cell_obj->sst_string);
///                     data_point->is_string = LXW_TRUE;
///                     range->has_string_cache = LXW_TRUE;
///                 }
///             }
///             else {
///                 data_point->no_data = LXW_TRUE;
///             }

///             STAILQ_INSERT_TAIL(range->data_cache, data_point, list_pointers);
///             num_data_points++;
///         }
///     }

///     range->num_data_points = num_data_points;
/// }

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
/// STATIC void _populate_range_dimensions(lxw_workbook *self, lxw_series_range *range)
/// {
///     char formula[LXW_MAX_FORMULA_RANGE_LENGTH] = { 0 };
///     char *tmp_str;
///     char *sheetname;

/* If neither the range formula or sheetname is defined then this probably
 * isn't a valid range.
 */
///     if (!range->formula && !range->sheetname) {
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }

/* If the sheetname is already defined it was already set via
 * chart_series_set_categories() or  chart_series_set_values().
 */
///     if (range->sheetname)
///         return;

/* Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5) */
///     if (range->formula[0] == '(') {
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }

/* Create a copy of the formula to modify and parse into parts. */
///     lxw_snprintf(formula, LXW_MAX_FORMULA_RANGE_LENGTH, "%s", range->formula);

/* Check for valid formula. Note, This needs stronger validation. */
///     tmp_str = strchr(formula, '!');

///     if (tmp_str == NULL) {
///         range->ignore_cache = LXW_TRUE;
///         return;
///     }
///     else {
/* Check for empty string. */
///         if (lxw_str_is_empty(tmp_str)) {
///             range->ignore_cache = LXW_TRUE;
///             return;
///         }

/* Split the formulas into sheetname and row-col data. */
///         *tmp_str = '\0';
///         tmp_str++;
///         sheetname = formula;

///         if (lxw_str_is_empty(tmp_str) || lxw_str_is_empty(sheetname)) {
///             range->ignore_cache = LXW_TRUE;
///             return;
///         }

/* Remove any worksheet quoting. */
///         if (sheetname[0] == '\'')
///             sheetname++;
///         if (strlen(sheetname) > 0 && sheetname[strlen(sheetname) - 1] == '\'') {
///             sheetname[strlen(sheetname) - 1] = '\0';
///         }

/* Check that the sheetname exists. */
///         if (!workbook_get_worksheet_by_name(self, sheetname)) {
///             LXW_WARN_FORMAT2("workbook_add_chart(): worksheet name '%s' "
///                              "in chart formula '%s' doesn't exist.",
///                              sheetname, range->formula);
///             range->ignore_cache = LXW_TRUE;
///             return;
///         }

///         range->sheetname = lxw_strdup(sheetname);
///         range->first_row = lxw_name_to_row(tmp_str);
///         range->first_col = lxw_name_to_col(tmp_str);

///         if (strchr(tmp_str, ':')) {
/* 2D range. */
///             range->last_row = lxw_name_to_row_2(tmp_str);
///             range->last_col = lxw_name_to_col_2(tmp_str);
///         }
///         else {
///             /* 1D range. */
///             range->last_row = range->first_row;
///             range->last_col = range->first_col;
///         }
///     }
/// }

/// STATIC void _populate_range(lxw_workbook *self, lxw_series_range *range)
/// {
///     if (!range)
///         return;

///     _populate_range_dimensions(self, range);
///     _populate_range_data_cache(self, range);
/// }

void workbook_t::add_chart_cache_data()
{
  ///     lxw_chart *chart;
  ///     lxw_chart_series *series;
  ///     uint16_t i;

  ///     STAILQ_FOREACH(chart, self->ordered_charts, ordered_list_pointers) {

  ///         _populate_range(self, chart->title.range);
  ///         _populate_range(self, chart->x_axis->title.range);
  ///         _populate_range(self, chart->y_axis->title.range);

  ///         if (STAILQ_EMPTY(chart->series_list))
  ///             continue;

  ///         STAILQ_FOREACH(series, chart->series_list, list_pointers) {
  ///             _populate_range(self, series->categories);
  ///             _populate_range(self, series->values);
  ///             _populate_range(self, series->title.range);

  ///             for (i = 0; i < series->data_label_count; i++) {
  ///                 lxw_chart_custom_label *data_label =
  ///                 &series->data_labels[i]; _populate_range(self,
  ///                 data_label->range);
  ///             }
  ///         }
  ///     }
}

/// STATIC void _store_image_type(lxw_workbook *self, uint8_t image_type)
/// {
///     if (image_type == LXW_IMAGE_PNG)
///         self->has_png = LXW_TRUE;

///     if (image_type == LXW_IMAGE_JPEG)
///         self->has_jpeg = LXW_TRUE;

///     if (image_type == LXW_IMAGE_BMP)
///         self->has_bmp = LXW_TRUE;

///     if (image_type == LXW_IMAGE_GIF)
///         self->has_gif = LXW_TRUE;
/// }

void workbook_t::prepare_drawings()
{
  ///     lxw_sheet *sheet;
  ///     lxw_worksheet *worksheet;
  ///     lxw_object_properties *object_props;
  ///     uint32_t chart_ref_id = 0;
  ///     uint32_t image_ref_id = 0;
  ///     uint32_t ref_id = 0;
  ///     uint32_t drawing_id = 0;
  ///     uint8_t is_chartsheet;
  ///     lxw_image_md5 tmp_image_md5;
  ///     lxw_image_md5 *new_image_md5 = NULL;
  ///     lxw_image_md5 *found_duplicate_image = NULL;
  ///     uint8_t i;

  ///     STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet) {
  ///             worksheet = sheet->u.chartsheet->worksheet;
  ///             is_chartsheet = LXW_TRUE;
  ///         }
  ///         else {
  ///             worksheet = sheet->u.worksheet;
  ///             is_chartsheet = LXW_FALSE;
  ///         }

  ///         if (STAILQ_EMPTY(worksheet->image_props)
  ///             && STAILQ_EMPTY(worksheet->embedded_image_props)
  ///             && STAILQ_EMPTY(worksheet->chart_data)
  ///             && !worksheet->has_header_vml &&
  ///             !worksheet->has_background_image) { continue;
  ///         }

  ///         drawing_id++;

  /* Prepare embedded worksheet images. */
  ///         STAILQ_FOREACH(object_props, worksheet->embedded_image_props,
  ///                        list_pointers) {

  ///             _store_image_type(self, object_props->image_type);

  /* Check for images with alt-text. */
  ///             if (object_props->description)
  ///                 self->has_embedded_image_descriptions = LXW_TRUE;

  /* Check for duplicate images and only store the first instance. */
  ///             if (object_props->md5) {
  ///                 tmp_image_md5.md5 = object_props->md5;
  ///                 found_duplicate_image = RB_FIND(lxw_image_md5s,
  ///                                                 self->embedded_image_md5s,
  ///                                                 &tmp_image_md5);
  ///             }

  ///             if (found_duplicate_image) {
  ///                 ref_id = found_duplicate_image->id;
  ///                 object_props->is_duplicate = LXW_TRUE;
  ///             }
  ///             else {
  ///                 image_ref_id++;
  ///                 ref_id = image_ref_id;
  ///                 self->num_embedded_images++;

  /// #ifndef USE_NO_MD5
  ///                 new_image_md5 = calloc(1, sizeof(lxw_image_md5));
  /// #endif
  ///                 if (new_image_md5 && object_props->md5) {
  ///                     new_image_md5->id = ref_id;
  ///                     new_image_md5->md5 = lxw_strdup(object_props->md5);

  ///                     RB_INSERT(lxw_image_md5s, self->embedded_image_md5s,
  ///                               new_image_md5);
  ///                 }
  ///             }

  ///             worksheet_set_error_cell(worksheet, object_props, ref_id);
  ///         }

  /* Prepare background images. */
  ///         if (worksheet->has_background_image) {

  ///             object_props = worksheet->background_image;

  ///             _store_image_type(self, object_props->image_type);

  /* Check for duplicate images and only store the first instance. */
  ///             if (object_props->md5) {
  ///                 tmp_image_md5.md5 = object_props->md5;
  ///                 found_duplicate_image = RB_FIND(lxw_image_md5s,
  ///                                                 self->background_md5s,
  ///                                                 &tmp_image_md5);
  ///             }

  ///             if (found_duplicate_image) {
  ///                 ref_id = found_duplicate_image->id;
  ///                 object_props->is_duplicate = LXW_TRUE;
  ///             }
  ///             else {
  ///                 image_ref_id++;
  ///                 ref_id = image_ref_id;

  /// #ifndef USE_NO_MD5
  ///                 new_image_md5 = calloc(1, sizeof(lxw_image_md5));
  /// #endif
  ///                 if (new_image_md5 && object_props->md5) {
  ///                     new_image_md5->id = ref_id;
  ///                     new_image_md5->md5 = lxw_strdup(object_props->md5);

  ///                     RB_INSERT(lxw_image_md5s, self->background_md5s,
  ///                               new_image_md5);
  ///                 }
  ///             }

  ///             lxw_worksheet_prepare_background(worksheet, ref_id,
  ///             object_props);
  ///         }

  /* Prepare worksheet images. */
  ///         STAILQ_FOREACH(object_props, worksheet->image_props,
  ///         list_pointers) {

  /* Ignore background image added above. */
  ///             if (object_props->is_background)
  ///                 continue;

  ///             _store_image_type(self, object_props->image_type);

  /* Check for duplicate images and only store the first instance. */
  ///             if (object_props->md5) {
  ///                 tmp_image_md5.md5 = object_props->md5;
  ///                 found_duplicate_image = RB_FIND(lxw_image_md5s,
  ///                                                 self->image_md5s,
  ///                                                 &tmp_image_md5);
  ///             }

  ///             if (found_duplicate_image) {
  ///                 ref_id = found_duplicate_image->id;
  ///                 object_props->is_duplicate = LXW_TRUE;
  ///             }
  ///             else {
  ///                 image_ref_id++;
  ///                 ref_id = image_ref_id;

  /// #ifndef USE_NO_MD5
  ///                 new_image_md5 = calloc(1, sizeof(lxw_image_md5));
  /// #endif
  ///                 if (new_image_md5 && object_props->md5) {
  ///                     new_image_md5->id = ref_id;
  ///                     new_image_md5->md5 = lxw_strdup(object_props->md5);

  ///                     RB_INSERT(lxw_image_md5s, self->image_md5s,
  ///                               new_image_md5);
  ///                 }
  ///             }

  ///             lxw_worksheet_prepare_image(worksheet, ref_id, drawing_id,
  ///                                         object_props);
  ///         }

  /* Prepare worksheet charts. */
  ///         STAILQ_FOREACH(object_props, worksheet->chart_data, list_pointers)
  ///         {
  ///             chart_ref_id++;
  ///             lxw_worksheet_prepare_chart(worksheet, chart_ref_id,
  ///             drawing_id,
  ///                                         object_props, is_chartsheet);
  ///             if (object_props->chart)
  ///                 STAILQ_INSERT_TAIL(self->ordered_charts,
  ///                 object_props->chart,
  ///                                    ordered_list_pointers);
  ///         }

  /* Prepare worksheet header/footer images. */
  ///         for (i = 0; i < LXW_HEADER_FOOTER_OBJS_MAX; i++) {

  ///             object_props = *worksheet->header_footer_objs[i];
  ///             if (!object_props)
  ///                 continue;

  ///             _store_image_type(self, object_props->image_type);

  /* Check for duplicate images and only store the first instance. */
  ///             if (object_props->md5) {
  ///                 tmp_image_md5.md5 = object_props->md5;
  ///                 found_duplicate_image = RB_FIND(lxw_image_md5s,
  ///                                                 self->header_image_md5s,
  ///                                                 &tmp_image_md5);
  ///             }

  ///             if (found_duplicate_image) {
  ///                 ref_id = found_duplicate_image->id;
  ///                 object_props->is_duplicate = LXW_TRUE;
  ///             }
  ///             else {
  ///                 image_ref_id++;
  ///                 ref_id = image_ref_id;

  /// #ifndef USE_NO_MD5
  ///                 new_image_md5 = calloc(1, sizeof(lxw_image_md5));
  /// #endif
  ///                 if (new_image_md5 && object_props->md5) {
  ///                     new_image_md5->id = ref_id;
  ///                     new_image_md5->md5 = lxw_strdup(object_props->md5);

  ///                     RB_INSERT(lxw_image_md5s, self->header_image_md5s,
  ///                               new_image_md5);
  ///                 }
  ///             }

  ///             lxw_worksheet_prepare_header_image(worksheet, ref_id,
  ///                                                object_props);
  ///         }

  ///     }

  ///     self->drawing_count = drawing_id;
}

void workbook_t::prepare_vml()
{
  ///     lxw_worksheet *worksheet;
  ///     lxw_sheet *sheet;
  ///     uint32_t comment_id = 0;
  ///     uint32_t vml_drawing_id = 0;
  ///     uint32_t vml_data_id = 1;
  ///     uint32_t vml_header_id = 0;
  ///     uint32_t vml_shape_id = 1024;
  ///     uint32_t comment_count = 0;

  ///     STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet)
  ///             continue;
  ///         else
  ///             worksheet = sheet->u.worksheet;

  ///         if (!worksheet->has_vml && !worksheet->has_header_vml)
  ///             continue;

  ///         if (worksheet->has_vml) {
  ///             self->has_vml = LXW_TRUE;
  ///             if (worksheet->has_comments) {
  ///                 self->comment_count++;
  ///                 comment_id++;
  ///                 self->has_comments = LXW_TRUE;
  ///             }

  ///             vml_drawing_id++;

  ///             comment_count = lxw_worksheet_prepare_vml_objects(worksheet,
  ///                                                               vml_data_id,
  ///                                                               vml_shape_id,
  ///                                                               vml_drawing_id,
  ///                                                               comment_id);

  /* Each VML should start with a shape id incremented by 1024. */
  ///             vml_data_id += 1 * ((1024 + comment_count) / 1024);
  ///             vml_shape_id += 1024 * ((1024 + comment_count) / 1024);
  ///         }

  ///         if (worksheet->has_header_vml) {
  ///             self->has_vml = LXW_TRUE;
  ///             vml_drawing_id++;
  ///             vml_header_id++;
  ///             lxw_worksheet_prepare_header_vml_objects(worksheet,
  ///                                                      vml_header_id,
  ///                                                      vml_drawing_id);
  ///         }
  ///     }
}

void workbook_t::prepare_defined_names()
{
  ///     lxw_worksheet *worksheet;
  ///     lxw_sheet *sheet;
  ///     char app_name[LXW_DEFINED_NAME_LENGTH];
  ///     char range[LXW_DEFINED_NAME_LENGTH];
  ///     char area[LXW_MAX_CELL_RANGE_LENGTH];
  ///     char first_col[8];
  ///     char last_col[8];

  ///     STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet)
  ///             continue;
  ///         else
  ///             worksheet = sheet->u.worksheet;
  /*
   * Check for autofilter settings and store them.
   */
  ///         if (worksheet->autofilter.in_use) {

  ///             lxw_snprintf(app_name, LXW_DEFINED_NAME_LENGTH,
  ///                          "%s!_FilterDatabase", worksheet->quoted_name);

  ///             lxw_rowcol_to_range_abs(area,
  ///                                     worksheet->autofilter.first_row,
  ///                                     worksheet->autofilter.first_col,
  ///                                     worksheet->autofilter.last_row,
  ///                                     worksheet->autofilter.last_col);

  ///             lxw_snprintf(range, LXW_DEFINED_NAME_LENGTH, "%s!%s",
  ///                          worksheet->quoted_name, area);

  /* Autofilters are the only defined name to set the hidden flag. */
  ///             _store_defined_name(self, "_xlnm._FilterDatabase", app_name,
  ///                                 range, worksheet->index, LXW_TRUE);
  ///         }

  /*
   * Check for Print Area settings and store them.
   */
  ///         if (worksheet->print_area.in_use) {

  ///             lxw_snprintf(app_name, LXW_DEFINED_NAME_LENGTH,
  ///                          "%s!Print_Area", worksheet->quoted_name);

  /* Check for print area that is the max row range. */
  ///             if (worksheet->print_area.first_row == 0
  ///                 && worksheet->print_area.last_row == LXW_ROW_MAX - 1) {

  ///                 lxw_col_to_name(first_col,
  ///                                 worksheet->print_area.first_col,
  ///                                 LXW_FALSE);

  ///                 lxw_col_to_name(last_col,
  ///                                 worksheet->print_area.last_col,
  ///                                 LXW_FALSE);

  ///                 lxw_snprintf(area, LXW_MAX_CELL_RANGE_LENGTH - 1,
  ///                 "$%s:$%s",
  ///                              first_col, last_col);

  ///             }
  /* Check for print area that is the max column range. */
  ///             else if (worksheet->print_area.first_col == 0
  ///                      && worksheet->print_area.last_col == LXW_COL_MAX - 1)
  ///                      {

  ///                 lxw_snprintf(area, LXW_MAX_CELL_RANGE_LENGTH - 1,
  ///                 "$%d:$%d",
  ///                              worksheet->print_area.first_row + 1,
  ///                              worksheet->print_area.last_row + 1);

  ///             }
  ///             else {
  ///                 lxw_rowcol_to_range_abs(area,
  ///                                         worksheet->print_area.first_row,
  ///                                         worksheet->print_area.first_col,
  ///                                         worksheet->print_area.last_row,
  ///                                         worksheet->print_area.last_col);
  ///             }

  ///             lxw_snprintf(range, LXW_DEFINED_NAME_LENGTH, "%s!%s",
  ///                          worksheet->quoted_name, area);

  ///             _store_defined_name(self, "_xlnm.Print_Area", app_name,
  ///                                 range, worksheet->index, LXW_FALSE);
  ///         }

  /*
   * Check for repeat rows/cols. aka, Print Titles and store them.
   */
  ///         if (worksheet->repeat_rows.in_use ||
  ///         worksheet->repeat_cols.in_use) {
  ///             if (worksheet->repeat_rows.in_use
  ///                 && worksheet->repeat_cols.in_use) {
  ///                 lxw_snprintf(app_name, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!Print_Titles", worksheet->quoted_name);

  ///                 lxw_col_to_name(first_col,
  ///                                 worksheet->repeat_cols.first_col,
  ///                                 LXW_FALSE);

  ///                 lxw_col_to_name(last_col,
  ///                                 worksheet->repeat_cols.last_col,
  ///                                 LXW_FALSE);

  ///                 lxw_snprintf(range, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!$%s:$%s,%s!$%d:$%d",
  ///                              worksheet->quoted_name, first_col,
  ///                              last_col, worksheet->quoted_name,
  ///                              worksheet->repeat_rows.first_row + 1,
  ///                              worksheet->repeat_rows.last_row + 1);

  ///                 _store_defined_name(self, "_xlnm.Print_Titles", app_name,
  ///                                     range, worksheet->index, LXW_FALSE);
  ///             }
  ///             else if (worksheet->repeat_rows.in_use) {

  ///                 lxw_snprintf(app_name, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!Print_Titles", worksheet->quoted_name);

  ///                 lxw_snprintf(range, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!$%d:$%d", worksheet->quoted_name,
  ///                              worksheet->repeat_rows.first_row + 1,
  ///                              worksheet->repeat_rows.last_row + 1);

  ///                 _store_defined_name(self, "_xlnm.Print_Titles", app_name,
  ///                                     range, worksheet->index, LXW_FALSE);
  ///             }
  ///             else if (worksheet->repeat_cols.in_use) {
  ///                 lxw_snprintf(app_name, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!Print_Titles", worksheet->quoted_name);

  ///                 lxw_col_to_name(first_col,
  ///                                 worksheet->repeat_cols.first_col,
  ///                                 LXW_FALSE);

  ///                 lxw_col_to_name(last_col,
  ///                                 worksheet->repeat_cols.last_col,
  ///                                 LXW_FALSE);

  ///                 lxw_snprintf(range, LXW_DEFINED_NAME_LENGTH,
  ///                              "%s!$%s:$%s", worksheet->quoted_name,
  ///                              first_col, last_col);

  ///                 _store_defined_name(self, "_xlnm.Print_Titles", app_name,
  ///                                     range, worksheet->index, LXW_FALSE);
  ///             }
  ///         }
  ///     }
}

void workbook_t::prepare_tables()
{
  ///     lxw_worksheet *worksheet;
  ///     lxw_sheet *sheet;
  ///     uint32_t table_id = 0;
  ///     uint32_t table_count = 0;

  ///     STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ///         if (sheet->is_chartsheet)
  ///             continue;
  ///         else
  ///             worksheet = sheet->u.worksheet;

  ///         table_count = worksheet->table_count;

  ///         if (table_count == 0)
  ///             continue;

  ///         lxw_worksheet_prepare_tables(worksheet, table_id + 1);

  ///         table_id += table_count;
  ///     }
}

std::string workbook_t::write_workbook() const
{
  return xml_start_tag("workbook", {
                                       {"xmlns",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main"      },
                                       {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/"
                                                   "relationships"},
  });
}

std::string workbook_t::write_file_version() const
{
  const std::vector<std::tuple<std::string, std::string>> attributes{
      {"appName",      "xl"  },
      {"lastEdited",   "4"   },
      {"lowestEdited", "4"   },
      {"rupBuild",     "4505"},
  };

  ///     if (self->vba_project)
  ///         LXW_PUSH_ATTRIBUTES_STR("codeName",
  ///                                 "{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}");

  return xml_empty_tag("fileVersion", attributes);
}

std::string workbook_t::write_file_sharing() const
{
  ///     if (self->read_only == 0)
  ///         return;

  ///     LXW_PUSH_ATTRIBUTES_STR("readOnlyRecommended", "1");
  ///     lxw_xml_empty_tag(self->file, "fileSharing", &attributes);

  return "";
}

std::string workbook_t::write_workbook_pr() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  ///     if (self->vba_codename)
  ///         LXW_PUSH_ATTRIBUTES_STR("codeName", self->vba_codename);

  ///     if (self->use_1904_epoch)
  ///         LXW_PUSH_ATTRIBUTES_STR("date1904", "1");

  attributes.emplace_back("defaultThemeVersion", "124226");

  return xml_empty_tag("workbookPr", attributes);
}

std::string workbook_t::write_workbook_view() const
{
  const std::vector<std::tuple<std::string, std::string>> attributes{
      {"xWindow",      "240"                         },
      {"yWindow",      "15"                          },
      {"windowWidth",  std::to_string(window_width_) },
      {"windowHeight", std::to_string(window_height_)},
  };

  ///     if (self->first_sheet)
  ///         LXW_PUSH_ATTRIBUTES_INT("firstSheet", self->first_sheet);

  ///     if (self->active_sheet)
  ///         LXW_PUSH_ATTRIBUTES_INT("activeTab", self->active_sheet);

  return xml_empty_tag("workbookView", attributes);
}

std::string workbook_t::write_book_views() const
{
  std::string xml_data = xml_start_tag("bookViews");
  xml_data += write_workbook_view();
  xml_data += xml_end_tag("bookViews");

  return xml_data;
}

std::string workbook_t::write_sheet(std::string_view name, uint32_t sheet_id/*,
             uint8_t hidden*/) const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"name", std::string(name)},
      {"sheetId", std::format("{}", sheet_id)},
  };

  ///     if (hidden)
  ///         LXW_PUSH_ATTRIBUTES_STR("state", "hidden");
  attributes.emplace_back("r:id", std::format("rId{}", sheet_id));

  return xml_empty_tag("sheet", attributes);
}

std::string workbook_t::write_sheets() const
{

  ///     lxw_sheet *sheet;
  ///     lxw_worksheet *worksheet;
  ///     lxw_chartsheet *chartsheet;

  std::string xml_data = xml_start_tag("sheets");

  for(auto sheet: sheets_)
  {
    ///         if (sheet->is_chartsheet) {
    ///             chartsheet = sheet->u.chartsheet;
    ///             _write_sheet(self, chartsheet->name, chartsheet->index + 1,
    ///                          chartsheet->hidden);
    ///         }
    ///         else {
    const auto ws = std::get<0>(sheet);
    xml_data += write_sheet(ws.get_sheet_name(), ws.get_sheet_index() + 1);

    ///             worksheet = sheet->u.worksheet;
    ///             _write_sheet(self, worksheet->name, worksheet->index + 1,
    ///                          worksheet->hidden);
    ///         }
  }

  xml_data += xml_end_tag("sheets");
  return xml_data;
}

std::string workbook_t::write_calc_pr() const
{
  return xml_empty_tag("calcPr", {
                                     {"calcId",         "124519"},
                                     {"fullCalcOnLoad", "1"     },
  });
}

/// STATIC void _write_defined_name(lxw_workbook *self, lxw_defined_name *defined_name)
/// {
///     struct xml_attribute_list attributes;
///     struct xml_attribute *attribute;

///     LXW_INIT_ATTRIBUTES();
///     LXW_PUSH_ATTRIBUTES_STR("name", defined_name->name);

///     if (defined_name->index != -1)
///         LXW_PUSH_ATTRIBUTES_INT("localSheetId", defined_name->index);

///     if (defined_name->hidden)
///         LXW_PUSH_ATTRIBUTES_INT("hidden", 1);

///     lxw_xml_data_element(self->file, "definedName", defined_name->formula,
///                          &attributes);

///     LXW_FREE_ATTRIBUTES();
/// }

std::string workbook_t::write_defined_names() const
{
  ///     lxw_defined_name *defined_name;

  ///     if (TAILQ_EMPTY(self->defined_names))
  ///         return;

  ///     lxw_xml_start_tag(self->file, "definedNames", NULL);

  ///     TAILQ_FOREACH(defined_name, self->defined_names, list_pointers) {
  ///         _write_defined_name(self, defined_name);
  ///     }

  ///     lxw_xml_end_tag(self->file, "definedNames");
  return "";
}

std::string workbook_t::assemble_xml_file()
{
  // Prepare workbook and sub-objects for writing.
  prepare_workbook();

  std::string xml_data = xml_declaration();
  xml_data += write_workbook();
  xml_data += write_file_version();
  xml_data += write_file_sharing();
  xml_data += write_workbook_pr();
  xml_data += write_book_views();
  xml_data += write_sheets();
  xml_data += write_defined_names();
  xml_data += write_calc_pr();
  xml_data += xml_end_tag("workbook");

  return xml_data;
}

workbook_t::workbook_t(/*lxw_workbook_options *options*/)
{
  ///     lxw_format *format;
  ///     lxw_workbook *workbook;

  /* Create the workbook object. */
  ///     workbook = calloc(1, sizeof(lxw_workbook));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook, mem_error);
  ///     workbook->filename = lxw_strdup(filename);

  /* Add the sheets list. */
  ///     workbook->sheets = calloc(1, sizeof(struct lxw_sheets));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->sheets, mem_error);
  ///     STAILQ_INIT(workbook->sheets);

  /* Add the worksheets list. */
  ///     workbook->worksheets = calloc(1, sizeof(struct lxw_worksheets));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->worksheets, mem_error);
  ///     STAILQ_INIT(workbook->worksheets);

  /* Add the chartsheets list. */
  ///     workbook->chartsheets = calloc(1, sizeof(struct lxw_chartsheets));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->chartsheets, mem_error);
  ///     STAILQ_INIT(workbook->chartsheets);

  /* Add the worksheet names tree. */
  ///     workbook->worksheet_names = calloc(1, sizeof(struct lxw_worksheet_names));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->worksheet_names, mem_error);
  ///     RB_INIT(workbook->worksheet_names);

  /* Add the chartsheet names tree. */
  ///     workbook->chartsheet_names = calloc(1,
  ///                                         sizeof(struct lxw_chartsheet_names));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->chartsheet_names, mem_error);
  ///     RB_INIT(workbook->chartsheet_names);

  /* Add the image MD5 tree. */
  ///     workbook->image_md5s = calloc(1, sizeof(struct lxw_image_md5s));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->image_md5s, mem_error);
  ///     RB_INIT(workbook->image_md5s);

  /* Add the embedded image MD5 tree. */
  ///     workbook->embedded_image_md5s = calloc(1, sizeof(struct lxw_image_md5s));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->embedded_image_md5s, mem_error);
  ///     RB_INIT(workbook->embedded_image_md5s);

  /* Add the header image MD5 tree. */
  ///     workbook->header_image_md5s = calloc(1, sizeof(struct lxw_image_md5s));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->header_image_md5s, mem_error);
  ///     RB_INIT(workbook->header_image_md5s);

  /* Add the background image MD5 tree. */
  ///     workbook->background_md5s = calloc(1, sizeof(struct lxw_image_md5s));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->background_md5s, mem_error);
  ///     RB_INIT(workbook->background_md5s);

  /* Add the charts list. */
  ///     workbook->charts = calloc(1, sizeof(struct lxw_charts));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->charts, mem_error);
  ///     STAILQ_INIT(workbook->charts);

  /* Add the ordered charts list to track chart insertion order. */
  ///     workbook->ordered_charts = calloc(1, sizeof(struct lxw_charts));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->ordered_charts, mem_error);
  ///     STAILQ_INIT(workbook->ordered_charts);

  /* Add the formats list. */
  ///     workbook->formats = calloc(1, sizeof(struct lxw_formats));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->formats, mem_error);
  ///     STAILQ_INIT(workbook->formats);

  /* Add the defined_names list. */
  ///     workbook->defined_names = calloc(1, sizeof(struct lxw_defined_names));
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->defined_names, mem_error);
  ///     TAILQ_INIT(workbook->defined_names);

  /* Add the shared strings table. */
  ///     workbook->sst = lxw_sst_new();
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->sst, mem_error);

  /* Add a hash table to track format indices. */
  ///     workbook->used_xf_formats = lxw_hash_new(128, 1, 0);
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->used_xf_formats, mem_error);

  /* Add a hash table to track format indices. */
  ///     workbook->used_dxf_formats = lxw_hash_new(128, 1, 0);
  ///     GOTO_LABEL_ON_MEM_ERROR(workbook->used_dxf_formats, mem_error);

  // Add the default cell format.
  auto format = add_format();
  // Initialize its index.
  get_xf_index(format);

  /* Add the default hyperlink format. */
  ///     format = workbook_add_format(workbook);
  ///     GOTO_LABEL_ON_MEM_ERROR(format, mem_error);
  ///     format_set_hyperlink(format);
  ///     workbook->default_url_format = format;

  ///     if (options) {
  ///         workbook->options.constant_memory = options->constant_memory;
  ///         workbook->options.tmpdir = lxw_strdup(options->tmpdir);
  ///         workbook->options.use_zip64 = options->use_zip64;
  ///         workbook->options.output_buffer = options->output_buffer;
  ///         workbook->options.output_buffer_size = options->output_buffer_size;
  ///     }
}

worksheet_t& workbook_t::add_worksheet()
{
  const std::string sheetname = "Sheet" + std::to_string(num_worksheets_ + 1);
  return add_worksheet(sheetname);
}

worksheet_t& workbook_t::add_worksheet(std::string_view sheetname)
{
  // Check that the worksheet name is valid.
  validate_sheetname(sheetname);

  ///     lxw_sheet *sheet = NULL;
  ///     lxw_worksheet *worksheet = NULL;
  ///     lxw_worksheet_name *worksheet_name = NULL;
  ///     lxw_error error;
  const worksheet_init_data_t init_data{
      .index_       = num_sheets_,
      .hidden_      = 0,
      ///     .optimize = self->options.constant_memory,
      ///     .active_sheet = &self->active_sheet,
      ///     .first_sheet = &self->first_sheet,
      .sst_         = &sst_,
      .name_        = std::string{sheetname},
      .quoted_name_ = quote_sheetname(sheetname),
      ///     .tmpdir = self->options.tmpdir,
      ///     .default_url_format = self->default_url_format,
      ///     .max_url_length = self->max_url_length,
      ///     .use_1904_epoch = self->use_1904_epoch,
  };

  sheets_.emplace_back(worksheet_t{init_data});
  num_worksheets_++;
  num_sheets_++;

  /* Create a struct to find/store the worksheet name/pointer. */
  ///     worksheet_name = calloc(1, sizeof(struct lxw_worksheet_name));
  ///     GOTO_LABEL_ON_MEM_ERROR(worksheet_name, mem_error);

  /* Store the worksheet so we can look it up by name. */
  ///     worksheet_name->name = init_data.name;
  ///     worksheet_name->worksheet = worksheet;
  ///     RB_INSERT(lxw_worksheet_names, self->worksheet_names, worksheet_name);

  return std::get<worksheet_t>(sheets_.back());
}

/// lxw_chartsheet * workbook_add_chartsheet(lxw_workbook *self, const char *sheetname)
/// {
///     lxw_sheet *sheet = NULL;
///     lxw_chartsheet *chartsheet = NULL;
///     lxw_chartsheet_name *chartsheet_name = NULL;
///     lxw_error error;
///     lxw_worksheet_init_data init_data =
///         { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
///     char *new_name = NULL;

///     if (sheetname) {
///         /* Use the user supplied name. */
///         init_data.name = lxw_strdup(sheetname);
///         init_data.quoted_name = lxw_quote_sheetname(sheetname);
///     }
///     else {
/* Use the default SheetN name. */
///         new_name = malloc(LXW_MAX_SHEETNAME_LENGTH);
///         GOTO_LABEL_ON_MEM_ERROR(new_name, mem_error);

///         lxw_snprintf(new_name, LXW_MAX_SHEETNAME_LENGTH, "Chart%d",
///                      self->num_chartsheets + 1);
///         init_data.name = new_name;
///         init_data.quoted_name = lxw_strdup(new_name);
///     }

/* Check that the chartsheet name is valid. */
///     error = workbook_validate_sheet_name(self, init_data.name);
///     if (error) {
///         LXW_WARN_FORMAT2
///             ("workbook_add_chartsheet(): chartsheet name '%s' has "
///              "error: %s", init_data.name, lxw_strerror(error));
///         goto mem_error;
///     }

/* Create a struct to find/store the chartsheet name/pointer. */
///     chartsheet_name = calloc(1, sizeof(struct lxw_chartsheet_name));
///     GOTO_LABEL_ON_MEM_ERROR(chartsheet_name, mem_error);

/* Initialize the metadata to pass to the chartsheet. */
///     init_data.hidden = 0;
///     init_data.index = self->num_sheets;
///     init_data.sst = self->sst;
///     init_data.optimize = self->options.constant_memory;
///     init_data.active_sheet = &self->active_sheet;
///     init_data.first_sheet = &self->first_sheet;
///     init_data.tmpdir = self->options.tmpdir;

/* Create a new chartsheet object. */
///     chartsheet = lxw_chartsheet_new(&init_data);
///     GOTO_LABEL_ON_MEM_ERROR(chartsheet, mem_error);

/* Add it to the chartsheet list. */
///     self->num_chartsheets++;
///     STAILQ_INSERT_TAIL(self->chartsheets, chartsheet, list_pointers);

/* Create a new sheet object. */
///     sheet = calloc(1, sizeof(lxw_sheet));
///     GOTO_LABEL_ON_MEM_ERROR(sheet, mem_error);
///     sheet->is_chartsheet = LXW_TRUE;
///     sheet->u.chartsheet = chartsheet;

/* Add it to the chartsheet list. */
///     self->num_sheets++;
///     STAILQ_INSERT_TAIL(self->sheets, sheet, list_pointers);

/* Store the chartsheet so we can look it up by name. */
///     chartsheet_name->name = init_data.name;
///     chartsheet_name->chartsheet = chartsheet;
///     RB_INSERT(lxw_chartsheet_names, self->chartsheet_names, chartsheet_name);

///     return chartsheet;

/// mem_error:
///     free((void *) init_data.name);
///     free((void *) init_data.quoted_name);
///     free(chartsheet_name);
///     free(chartsheet);
///     return NULL;
/// }

/// lxw_chart * workbook_add_chart(lxw_workbook *self, uint8_t type)
/// {
///     lxw_chart *chart;

///     if (type == LXW_CHART_NONE || type > LXW_CHART_RADAR_FILLED) {
///         LXW_WARN_FORMAT1("workbook_add_chart(): invalid chart type: %d",
///                          type);
///         return NULL;
///     }

/* Create a new chart object. */
///     chart = lxw_chart_new(type);

///     if (chart)
///         STAILQ_INSERT_TAIL(self->charts, chart, list_pointers);

///     return chart;
/// }

// TODO Add class that encapsulate this pointer for interaction with caller. Pointers will only be used inside library.
// TODO Constructor and pointer of this class should be only usable by workbook and worksheet (friendship)
format_t* workbook_t::add_format()
{
  formats_.emplace_back();

  return &formats_.back();
}

int32_t workbook_t::get_xf_index(format_t* format)
{
  // Format already has an index number so return it.
  if(format->xf_index_ != format_t::PROPERTY_UNSET)
  {
    return format->xf_index_;
  }

  // TODO Add search two be sure there is no duplication.
  format->xf_index_ = static_cast<int32_t>(used_xf_formats_.size());
  used_xf_formats_.push_back(format);

  return format->xf_index_;
}

void workbook_t::save(std::string_view filename)
{
  ////   lxw_sheet *sheet = NULL;
  ////    lxw_worksheet *worksheet = NULL;
  ////    lxw_packager *packager = NULL;
  ////    lxw_error error = LXW_NO_ERROR;
  ////    char codename[LXW_MAX_SHEETNAME_LENGTH] = { 0 };

  // Add a default worksheet if non have been added.
  if(num_sheets_ == 0)
  {
    add_worksheet();
  }

  // Ensure that at least one worksheet has been selected.
  if(active_sheet_ == 0)
  {
    auto& sheet = std::get<worksheet_t>(sheets_.front());
    ////  TODO if (!sheet->is_chartsheet) {
    sheet.select();
    ////        }
  }

  /* Set the active sheet and check if a metadata file is needed. */
  ////    STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ////        if (sheet->is_chartsheet)
  ////            continue;
  ////        else
  ////            worksheet = sheet->u.worksheet;

  ////        if (worksheet->index == self->active_sheet)
  ////            worksheet->active = LXW_TRUE;

  ////        if (worksheet->has_dynamic_functions) {
  ////            self->has_metadata = LXW_TRUE;
  ////            self->has_dynamic_functions = LXW_TRUE;
  ////        }

  ////        if (!STAILQ_EMPTY(worksheet->embedded_image_props)) {
  ////            self->has_metadata = LXW_TRUE;
  ////            self->has_embedded_images = LXW_TRUE;
  ////        }
  ////}

  /* Set workbook and worksheet VBA codenames if a macro has been added. */
  ////    if (self->vba_project) {
  ////        if (!self->vba_codename)
  ////            workbook_set_vba_name(self, "ThisWorkbook");

  ////        STAILQ_FOREACH(sheet, self->sheets, list_pointers) {
  ////            if (sheet->is_chartsheet)
  ////                continue;
  ////            else
  ////                worksheet = sheet->u.worksheet;

  ////            if (!worksheet->vba_codename) {
  ////                lxw_snprintf(codename, LXW_MAX_SHEETNAME_LENGTH,
  ///"Sheet%d", /                             worksheet->index + 1);

  ////                worksheet_set_vba_name(worksheet, codename);
  ////            }
  ////        }
  ////    }

  // Prepare the worksheet VML elements such as comments.
  prepare_vml();

  // Set the defined names for the worksheets such as Print Titles.
  prepare_defined_names();

  // Prepare the drawings, charts and images.
  prepare_drawings();

  // Add cached data to charts.
  add_chart_cache_data();

  // Set the table ids for the worksheet tables.
  prepare_tables();

  // Create a packager object to assemble sub-elements into a zip file.
  packager_t pkg(filename);
  ////    packager = lxw_packager_new(self->filename,
  ////                                self->options.tmpdir,
  ////                                self->options.use_zip64);

  pkg.create_package(*this);

  ////    if (!self->filename) {
  ////        *self->options.output_buffer = packager->output_buffer;
  ////        *self->options.output_buffer_size = packager->output_buffer_size;
  ////    }

  /* Error and non-error conditions fall through to the cleanup code. */
  ////    if (error == LXW_ERROR_CREATING_TMPFILE) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Error creating tmpfile(s) to assemble '%s'. "
  ////                   "System error = %s\n", self->filename,
  /// strerror(errno)); /    }

  /* If LXW_ERROR_ZIP_FILE_OPERATION then errno is set by zip. */
  ////    if (error == LXW_ERROR_ZIP_FILE_OPERATION) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip ZIP_ERRNO error while creating xlsx file '%s'. "
  ////                   "System error = %s\n", self->filename,
  /// strerror(errno)); /    }

  /* If LXW_ERROR_ZIP_PARAMETER_ERROR then errno is set by zip. */
  ////    if (error == LXW_ERROR_ZIP_PARAMETER_ERROR) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip ZIP_PARAMERROR error while creating xlsx file
  ///'%s'. " /                   "System error = %s\n", self->filename,
  /// strerror(errno)); /    }

  /* If LXW_ERROR_ZIP_BAD_ZIP_FILE then errno is set by zip. */
  ////    if (error == LXW_ERROR_ZIP_BAD_ZIP_FILE) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip ZIP_BADZIPFILE error while creating xlsx file
  ///'%s'. " /                   "This may require the use_zip64 option for
  /// large files. " /                   "System error = %s\n", self->filename,
  /// strerror(errno)); /    }

  /* If LXW_ERROR_ZIP_INTERNAL_ERROR then errno is set by zip. */
  ////    if (error == LXW_ERROR_ZIP_INTERNAL_ERROR) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip ZIP_INTERNALERROR error while creating xlsx file
  ///'%s'. " /                   "System error = %s\n", self->filename,
  /// strerror(errno)); /    }

  /* The next 2 error conditions don't set errno. */
  ////    if (error == LXW_ERROR_ZIP_FILE_ADD) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip error adding file to xlsx file '%s'.\n",
  ////                   self->filename);
  ////    }

  ////    if (error == LXW_ERROR_ZIP_CLOSE) {
  ////        LXW_PRINTF(LXW_STDERR "[ERROR] workbook_close(): "
  ////                   "Zip error closing xlsx file '%s'.\n", self->filename);
  ////    }

  ////mem_error:
  ////    lxw_packager_free(packager);
  ////    lxw_workbook_free(self);
  ////    return error;
}

/// lxw_error workbook_define_name(lxw_workbook *self, const char *name,
///                      const char *formula)
/// {
///     return _store_defined_name(self, name, NULL, formula, -1, LXW_FALSE);
/// }

void workbook_t::set_properties(const doc_properties_t& properties)
{
  properties_ = properties;
}

void workbook_t::set_custom_property(std::string_view name, const std::string& value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("Name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Name of custom property ({}) is too long", name));
  }

  if(value.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Value of custom property ({}) is too long", value));
  }

  custom_properties_.emplace_back(custom_property_types_t::STRING, std::string(name), std::string(value));
}

void workbook_t::set_custom_property(std::string_view name, const char* value)
{
  set_custom_property(name, std::string(value));
}

void workbook_t::set_custom_property(std::string_view name, int32_t value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("Name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Name of custom property ({}) is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::INTEGER, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, double value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("Name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Name of custom property ({}) is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::DOUBLE, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, bool value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("Name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Name of custom property ({}) is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::BOOLEAN, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::system_clock::time_point& value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("Name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(std::format("Name of custom property ({}) is too long", name));
  }

  if(value.time_since_epoch().count() == 0)
  {
    throw xwpp_exception_t("No date set");
  }

  custom_properties_.emplace_back(custom_property_types_t::DATETIME, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::year_month_day& value)
{
  set_custom_property(name, std::chrono::sys_days{value} + 0h + 0min + 0s + 0ms);
}

/// lxw_worksheet * workbook_get_worksheet_by_name(lxw_workbook *self, const char *name)
/// {
///     lxw_worksheet_name worksheet_name;
///     lxw_worksheet_name *found;

///     if (!name)
///         return NULL;

///     worksheet_name.name = name;
///     found = RB_FIND(lxw_worksheet_names,
///                     self->worksheet_names, &worksheet_name);

///     if (found)
///         return found->worksheet;
///     else
///         return NULL;
/// }

/// lxw_chartsheet * workbook_get_chartsheet_by_name(lxw_workbook *self, const char *name)
/// {
///     lxw_chartsheet_name chartsheet_name;
///     lxw_chartsheet_name *found;

///     if (!name)
///         return NULL;

///     chartsheet_name.name = name;
///     found = RB_FIND(lxw_chartsheet_names,
///                     self->chartsheet_names, &chartsheet_name);

///     if (found)
///         return found->chartsheet;
///     else
///         return NULL;
/// }

/// lxw_format * workbook_get_default_url_format(lxw_workbook *self)
/// {
///     return self->default_url_format;
/// }

/// void workbook_unset_default_url_format(lxw_workbook *self)
/// {
///     self->default_url_format->hyperlink = LXW_FALSE;
///     self->default_url_format->xf_id = 0;
///     self->default_url_format->underline = LXW_UNDERLINE_NONE;
///     self->default_url_format->theme = 0;
/// }

void workbook_t::validate_sheetname(std::string_view sheetname) const
{
  // Check for empty worksheet name.
  if(sheetname.empty())
  {
    throw xwpp_exception_t("Sheetname is empty");
  }

  // Check the length of the worksheet name.
  if(sheetname.size() > XWPP_SHEETNAME_MAX)
  {
    throw xwpp_exception_t(std::format("Sheetname '{}' is too long", sheetname));
  }

  // Check that the worksheet name doesn't contain invalid characters.
  if(sheetname.find_first_of("[]:*?/\\") != std::string::npos)
  {
    throw xwpp_exception_t(std::format("Sheetname '{}' contains invalid characters", sheetname));
  }

  // Check that the worksheet doesn't start or end with an apostrophe.
  if(sheetname[0] == '\'' || sheetname[sheetname.size() - 1] == '\'')
  {
    throw xwpp_exception_t(std::format("Sheetname '{}' contains unbalanced single quote", sheetname));
  }

  // Check if the worksheet name is already in use.
  ///     if (workbook_get_worksheet_by_name(self, sheetname))
  ///         return LXW_ERROR_SHEETNAME_ALREADY_USED;

  // Check if the chartsheet name is already in use.
  ///     if (workbook_get_chartsheet_by_name(self, sheetname))
  ///         return LXW_ERROR_SHEETNAME_ALREADY_USED;
}

/// lxw_error workbook_add_vba_project(lxw_workbook *self, const char *filename)
/// {
///     FILE *filehandle;

///     if (!filename) {
///         LXW_WARN("workbook_add_vba_project(): "
///                  "project filename must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }

/* Check that the vbaProject file exists and can be opened. */
///     filehandle = lxw_fopen(filename, "rb");
///     if (!filehandle) {
///         LXW_WARN_FORMAT1("workbook_add_vba_project(): "
///                          "project file doesn't exist or can't be opened: %s.",
///                          filename);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     fclose(filehandle);

///     self->vba_project = lxw_strdup(filename);

///     return LXW_NO_ERROR;
/// }

/// lxw_error workbook_add_signed_vba_project(lxw_workbook *self,
///                                 const char *vba_project,
///                                 const char *signature)
/// {
///     FILE *filehandle;

///     lxw_error error = workbook_add_vba_project(self, vba_project);
///     if (error != LXW_NO_ERROR)
///         return error;

///     if (!signature) {
///         LXW_WARN("workbook_add_signed_vba_project(): "
///                  "signature filename must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }

/* Check that the vbaProjectSignature file exists and can be opened. */
///     filehandle = lxw_fopen(signature, "rb");
///     if (!filehandle) {
///         LXW_WARN_FORMAT1("workbook_add_signed_vba_project(): "
///                          "signature file doesn't exist or can't be opened: %s.",
///                          signature);
///         return LXW_ERROR_PARAMETER_VALIDATION;
///     }
///     fclose(filehandle);

///     self->vba_project_signature = lxw_strdup(signature);

///     return LXW_NO_ERROR;
/// }

/// lxw_error workbook_set_vba_name(lxw_workbook *self, const char *name)
/// {
///     if (!name) {
///         LXW_WARN("workbook_set_vba_name(): " "name must be specified.");
///         return LXW_ERROR_NULL_PARAMETER_IGNORED;
///     }

///     if (lxw_str_is_empty(name)) {
///         LXW_WARN_FORMAT("workbook_set_vba_name(): parameter "
///                         "'name' cannot be an empty string.");
///         return LXW_ERROR_PARAMETER_IS_EMPTY;
///     }

///     self->vba_codename = lxw_strdup(name);

///     return LXW_NO_ERROR;
/// }

/// void workbook_read_only_recommended(lxw_workbook *self)
/// {
///     self->read_only = 2;
/// }

/// void workbook_use_1904_epoch(lxw_workbook *self)
/// {
///     self->use_1904_epoch = LXW_TRUE;
/// }

/// void workbook_set_size(lxw_workbook *workbook, uint16_t width, uint16_t height)
/// {
/* Convert the width/height to twips at 96 dpi. */
///     if (width)
///         workbook->window_width = width * 1440 / 96;

///     if (height)
///         workbook->window_height = height * 1440 / 96;
/// }

}
