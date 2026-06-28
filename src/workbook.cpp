/*
 * workbook - A library for creating Excel XLSX workbook files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/workbook.h"

#include "xwpp/exception.h"
#include "xwpp/packager.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <algorithm>
#include <cctype>
#include <chrono>
#include <format>
#include <functional>
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

void workbook_t::prepare_fonts()
{
  std::vector<format_t*> fonts;

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

  // For DXF formats we only need to check if the properties have changed.
  for(auto format: used_dxf_formats_)
  {
    // The only font properties that can change for a DXF format are:
    // color, bold, italic, underline and strikethrough.
    if(format->font_color_ != color_t::UNSET || format->bold_ || format->italic_ ||
       format->underline_ != format_underlines_t::NONE || format->font_strikeout_)
    {
      format->has_dxf_font_ = true;
    }
  }

  font_count_ = static_cast<uint32_t>(fonts.size());
}

void workbook_t::prepare_borders()
{
  std::vector<format_t*> borders;

  // TODO Use unordered_set to optimise this search
  for(const auto format: used_xf_formats_)
  {
    for(const auto border: borders)
    {
      if(format->bottom_ == border->bottom_ && format->left_ == border->left_ && format->right_ == border->right_ &&
         format->top_ == border->top_ && format->diag_border_ == border->diag_border_ &&
         format->diag_type_ == border->diag_type_ && format->bottom_color_ == border->bottom_color_ &&
         format->left_color_ == border->left_color_ && format->right_color_ == border->right_color_ &&
         format->top_color_ == border->top_color_ && format->diag_color_ == border->diag_color_)
      {
        // Font has already been used.
        format->border_index_ = border->border_index_;
        format->has_border_   = false;
      }
    }

    if(format->border_index_ == format_t::PROPERTY_UNSET)
    {
      format->border_index_ = static_cast<int32_t>(borders.size());
      format->has_border_   = true;
      borders.push_back(format);
    }
  }

  /* For DXF formats we only need to check if the properties have changed. */
  ///     LXW_FOREACH_ORDERED(used_format_element, self->used_dxf_formats) {
  ///         lxw_format *format = (lxw_format *) used_format_element->value;

  ///         if (format->left || format->right || format->top || format->bottom) {
  ///             format->has_dxf_border = LXW_TRUE;
  ///         }
  ///     }

  border_count_ = static_cast<uint32_t>(borders.size());
}

void workbook_t::prepare_fills()
{
  std::vector<format_t*> fills;

  // Add the default fills.
  format_t default_fill_1;
  default_fill_1.fg_color_   = color_t::UNSET;
  default_fill_1.bg_color_   = color_t::UNSET;
  default_fill_1.pattern_    = format_patterns_t::NONE;
  default_fill_1.fill_index_ = 0;
  fills.push_back(&default_fill_1);
  format_t default_fill_2;
  default_fill_2.fg_color_   = color_t::UNSET;
  default_fill_2.bg_color_   = color_t::UNSET;
  default_fill_2.pattern_    = format_patterns_t::GRAY_125;
  default_fill_2.fill_index_ = 1;
  fills.push_back(&default_fill_2);

  // For DXF formats we only need to check if the properties have changed.
  for(auto format: used_dxf_formats_)
  {
    if(format->pattern_ != format_patterns_t::NONE || format->bg_color_ != color_t::UNSET ||
       format->fg_color_ != color_t::UNSET)
    {
      format->has_dxf_fill_ = true;
      format->dxf_bg_color_ = format->bg_color_;
      format->dxf_fg_color_ = format->fg_color_;
    }
  }

  // TODO Use unordered_set to optimise this search
  for(auto format: used_xf_formats_)
  {
    /* The following logical statements jointly take care of special */
    /* cases in relation to cell colors and patterns:                */
    /* 1. For a solid fill (pattern == 1) Excel reverses the role of */
    /*    foreground and background colors, and                      */
    /* 2. If the user specifies a foreground or background color     */
    /*    without a pattern they probably wanted a solid fill, so    */
    /*    we fill in the defaults.
     */
    if(format->pattern_ == format_patterns_t::SOLID && format->bg_color_ != color_t::UNSET &&
       format->fg_color_ != color_t::UNSET)
    {
      std::swap(format->bg_color_, format->fg_color_);
    }
    else if((format->pattern_ == format_patterns_t::SOLID || format->pattern_ == format_patterns_t::NONE) &&
            format->bg_color_ != color_t::UNSET && format->fg_color_ == color_t::UNSET)
    {
      std::swap(format->bg_color_, format->fg_color_);
      format->pattern_ = format_patterns_t::SOLID;
    }
    else if((format->pattern_ == format_patterns_t::SOLID || format->pattern_ == format_patterns_t::NONE) &&
            format->bg_color_ == color_t::UNSET && format->fg_color_ != color_t::UNSET)
    {
      format->pattern_ = format_patterns_t::SOLID;
    }

    for(const auto fill: fills)
    {
      if(format->bg_color_ == fill->bg_color_ && format->fg_color_ == fill->fg_color_ &&
         format->pattern_ == fill->pattern_)
      {
        // Font has already been used.
        format->fill_index_ = fill->fill_index_;
        format->has_fill_   = false;
      }
    }

    if(format->fill_index_ == format_t::PROPERTY_UNSET)
    {
      format->fill_index_ = static_cast<int32_t>(fills.size());
      format->has_fill_   = true;
      fills.push_back(format);
    }
  }

  fill_count_ = static_cast<uint32_t>(fills.size());
  ;
}

void workbook_t::prepare_num_formats()
{
  std::vector<format_t*> num_formats;

  ///     lxw_hash_table *num_formats = lxw_hash_new(128, 0, 1);
  ///     lxw_hash_element *hash_element;
  ///     lxw_hash_element *used_format_element;
  ///     uint16_t index = 0xA4;
  ///     uint16_t num_format_count = 0;
  ///     uint16_t *num_format_index;

  // TODO Use unordered_set to optimise this search
  for(const auto format: used_xf_formats_)
  {
    // Format already has a number format index.
    if(format->num_format_index_ != 0)
    {
      continue;
    }

    // Check if there is a user defined number format string.
    if(!format->num_format_.empty())
    {
      for(const auto num_format: num_formats)
      {
        if(format->num_format_ == num_format->num_format_)
        {
          // Format number has already been used.
          format->num_format_index_ = num_format->num_format_index_;
        }
      }

      if(format->num_format_index_ == 0)
      {
        // Custom number formats start at 0xA4
        format->num_format_index_ = static_cast<int32_t>(num_formats.size()) + 0xA4;
        num_formats.push_back(format);
      }
    }
  }

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

  num_format_count_ = static_cast<uint32_t>(num_formats.size());
}

void workbook_t::prepare_workbook()
{
  prepare_fonts();
  prepare_num_formats();
  prepare_borders();
  prepare_fills();
}

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
void workbook_t::store_defined_name(const std::string& name, const std::string& app_name, const std::string& formula,
                                    int16_t index, bool hidden)
{
  // Do some checks on the input data
  if(name.empty() || formula.empty())
  {
    throw xwpp_exception_t("workbook_t::store_defined_name(): 'name' and 'formula' cannot be empty");
  }

  defined_name_t defined_name;

  // Set the worksheet index or -1 for a global defined name.
  defined_name.index_  = index;
  defined_name.hidden_ = hidden;

  // Check for local defined names like like "Sheet1!name".
  size_t found_string = name.find('!');
  if(found_string == std::string::npos)
  {
    // The name is global. We just store the defined name string.
    defined_name.name_ = name;
  }
  else
  {
    // The name is worksheet local. We need to extract the sheet name
    // and map it to a sheet index.
    std::string worksheet_name = name.substr(0, found_string);
    std::string tmp_str        = name.substr(found_string + 1);

    if(tmp_str.empty() || worksheet_name.empty())
    {
      throw xwpp_exception_t("workbook_t::store_defined_name(): nor sheetname neither area can be empty.");
    }

    // Remove any worksheet quoting.
    if(worksheet_name.front() == '\'')
    {
      worksheet_name = worksheet_name.substr(1);
    }
    if(worksheet_name.back() == '\'')
    {
      worksheet_name.pop_back();
    }

    // Search for worksheet name to get the equivalent worksheet index.
    for(auto& sheet: sheets_)
    {
      if(std::holds_alternative<worksheet_t>(sheet))
      {
        auto& ws = std::get<worksheet_t>(sheet);

        if(worksheet_name == ws.name_)
        {
          defined_name.index_                = ws.index_;
          defined_name.normalised_sheetname_ = worksheet_name;
        }
      }
    }

    // If we didn't find the worksheet name we exit.
    if(defined_name.index_ == -1)
    {
      throw xwpp_exception_t(
          std::format("workbook_t::store_defined_name(): Sheename '{}' not present.", worksheet_name));
    }

    defined_name.name_ = tmp_str;
  }

  // Print titles and repeat title pass in the name used for App.xml.
  if(!app_name.empty())
  {
    defined_name.app_name_             = app_name;
    defined_name.normalised_sheetname_ = app_name;
  }
  else
  {
    defined_name.app_name_ = name;
  }

  // We need to normalize the defined names for sorting. This involves
  // removing any _xlnm namespace  and converting it to lowercase.
  found_string = defined_name.name_.find("_xlnm.");
  if(found_string == std::string::npos)
  {
    defined_name.normalised_name_ = defined_name.name_;
  }
  else
  {
    defined_name.normalised_name_ = defined_name.name_.substr(found_string + 6);
  }

  std::transform(std::begin(defined_name.normalised_name_), std::end(defined_name.normalised_name_),
                 std::begin(defined_name.normalised_name_), [](char c) { return std::tolower(c); });
  std::transform(std::begin(defined_name.normalised_sheetname_), std::end(defined_name.normalised_sheetname_),
                 std::begin(defined_name.normalised_sheetname_), [](char c) { return std::tolower(c); });

  // Strip leading "=" from the formula.
  if(formula[0] == '=')
  {
    defined_name.formula_ = formula.substr(1);
  }
  else
  {
    defined_name.formula_ = formula;
  }

  // We add the defined name to the list in sorted order.
  for(auto it = std::cbegin(defined_names_); it != std::cend(defined_names_); ++it)
  {
    if(defined_name.normalised_name_ == it->normalised_name_ &&
       defined_name.normalised_sheetname_ == it->normalised_sheetname_)
    {
      // The entry already exists. We exit and don't overwrite.
      return;
    }
    else if(defined_name.normalised_name_ < it->normalised_name_ ||
            (defined_name.normalised_name_ == it->normalised_name_ &&
             defined_name.normalised_sheetname_ < it->normalised_sheetname_))
    {
      // New defined name is inserted in sorted order before other entries.
      defined_names_.insert(it, defined_name);
      return;
    }
  }

  // Insert at the end otherwise
  defined_names_.push_back(defined_name);
}

/*
 * Populate the data cache of a chart data series by reading the data from the
 * relevant worksheet and adding it to the cached in the range object as a
 * list of points.
 *
 * Note, the data cache isn't strictly required by Excel but it helps if the
 * chart is embedded in another application such as PowerPoint and it also
 * helps with comparison testing.
 */
void workbook_t::populate_range_data_cache(series_range_t& range)
{
  ///     lxw_worksheet *worksheet;
  ///     row_num_t row_num;
  ///     col_num_t col_num;
  ///     row_t *row_obj;
  ///     cell_t *cell_obj;
  ///     struct lxw_series_data_point *data_point;
  uint16_t num_data_points = 0;

  // If ignore_cache is set then don't try to populate the cache. This flag
  // may be set manually, for testing, or due to a case where the cache can't be calculated.
  if(range.ignore_cache_)
  {
    return;
  }

  // Currently we only handle 2D ranges so ensure either the rows or cols are the same.
  if(range.first_row_ != range.last_row_ && range.first_col_ != range.last_col_)
  {
    range.ignore_cache_ = true;
    return;
  }

  // Check that the sheetname exists.
  const worksheet_t* worksheet = get_worksheet_by_name(range.sheetname_);
  if(worksheet == nullptr)
  {
    range.ignore_cache_ = true;
    throw xwpp_exception_t(
        std::format("workbook_t::populate_range_data_cache(): sheetname '{}' in chart formula '{}' doesn't exist",
                    range.sheetname_, range.formula_));
  }

  /* We can't read the data when worksheet optimization is on. */
  ///     if (worksheet->optimize) {
  ///         range->ignore_cache = LXW_TRUE;
  ///         return;
  ///     }

  // Iterate through the worksheet data and populate the range cache.
  for(row_num_t row_num = range.first_row_; row_num <= range.last_row_; row_num++)
  {
    const row_t* row_obj = worksheet->find_row(row_num);
    for(col_num_t col_num = range.first_col_; col_num <= range.last_col_; col_num++)
    {
      series_data_point_t data_point;
      const cell_t* cell_obj = worksheet->find_cell_in_row(row_obj, col_num);

      if(cell_obj)
      {
        if(cell_obj->type_ == cell_types_t::NUMBER_CELL)
        {
          data_point.number_ = std::get<double>(cell_obj->data_);
        }

        if(cell_obj->type_ == cell_types_t::STRING_CELL)
        {
          data_point.str_         = cell_obj->sst_string_;
          data_point.is_string_   = true;
          range.has_string_cache_ = true;
        }
      }
      else
      {
        data_point.no_data_ = true;
      }

      range.data_cache_.push_back(data_point);
      num_data_points++;
    }
  }
  range.num_data_points_ = num_data_points;
}

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
void workbook_t::populate_range_dimensions(series_range_t& range)
{
  // If neither the range formula or sheetname is defined then this probably
  // isn't a valid range.
  if(range.formula_.empty() && range.sheetname_.empty())
  {
    range.ignore_cache_ = true;
    return;
  }

  // If the sheetname is already defined it was already set via
  // chart_series_set_categories() or  chart_series_set_values().
  if(!range.sheetname_.empty())
  {
    return;
  }

  // Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5)
  if(range.formula_[0] == '(')
  {
    range.ignore_cache_ = true;
    return;
  }

  // Create a copy of the formula to modify and parse into parts.
  std::string formula = range.formula_;

  // Check for valid formula. Note, This needs stronger validation.
  size_t found_string = formula.find('!');
  if(found_string == std::string::npos)
  {
    range.ignore_cache_ = true;
    return;
  }
  else
  {
    // Split the formulas into sheetname and row-col data.
    std::string sheetname = formula.substr(0, found_string);
    std::string tmp_str   = formula.substr(found_string + 1);

    if(tmp_str.empty() || sheetname.empty())
    {
      range.ignore_cache_ = true;
      return;
    }

    // Remove any worksheet quoting.
    if(sheetname.front() == '\'')
    {
      sheetname = sheetname.substr(1);
    }
    if(sheetname.back() == '\'')
    {
      sheetname.pop_back();
    }

    // Check that the sheetname exists.
    if(get_worksheet_by_name(sheetname) == nullptr)
    {
      throw xwpp_exception_t(
          std::format("workbook_t::populate_range_dimensions(): sheetname '{}' in chart formula '{}' doesn't exist",
                      sheetname, range.formula_));
    }

    range.sheetname_ = sheetname;
    range.first_row_ = name_to_row(tmp_str.c_str());
    range.first_col_ = name_to_col(tmp_str.c_str());

    size_t found_string = formula.find(':');
    if(found_string == std::string::npos)
    {
      // 1D range.
      range.last_row_ = range.first_row_;
      range.last_col_ = range.first_col_;
    }
    else
    {
      // 2D range.
      range.last_row_ = name_to_row_2(tmp_str.c_str());
      range.last_col_ = name_to_col_2(tmp_str.c_str());
    }
  }
}

void workbook_t::populate_range(series_range_t& range)
{
  populate_range_dimensions(range);
  populate_range_data_cache(range);
}

void workbook_t::add_chart_cache_data()
{
  for(auto chart: ordered_charts_)
  {
    populate_range(chart->title_.range_);
    populate_range(chart->x_axis_.title_.range_);
    populate_range(chart->y_axis_.title_.range_);

    for(auto& series: chart->series_list_)
    {
      populate_range(series.categories_);
      populate_range(series.values_);
      populate_range(series.title_.range_);
      for(auto& data_label: series.data_labels_)
      {
        if(data_label.range_)
        {
          populate_range(*data_label.range_);
        }
      }
    }
  }
}

void workbook_t::store_image_type(image_types_t image_type)
{
  if(image_type == image_types_t::PNG)
  {
    has_png_ = true;
  }

  if(image_type == image_types_t::JPEG)
  {
    has_jpeg_ = true;
  }

  if(image_type == image_types_t::BMP)
  {
    has_bmp_ = true;
  }

  if(image_type == image_types_t::GIF)
  {
    has_gif_ = true;
  }
}

void workbook_t::prepare_drawings()
{
  ///     lxw_sheet *sheet;
  ///     lxw_object_properties *object_props;
  uint32_t chart_ref_id = 0;
  uint32_t image_ref_id = 0;
  ///     uint32_t ref_id = 0;
  uint32_t drawing_id   = 0;
  bool is_chartsheet;
  ///     lxw_image_md5 tmp_image_md5;
  ///     lxw_image_md5 *new_image_md5 = NULL;
  ///     lxw_image_md5 *found_duplicate_image = NULL;
  ///     uint8_t i;

  for(auto& sheet: sheets_)
  {
    auto& worksheet = std::holds_alternative<worksheet_t>(sheet) ? std::get<worksheet_t>(sheet)
                                                                 : std::get<chartsheet_t>(sheet).worksheet_;
    is_chartsheet   = std::holds_alternative<chartsheet_t>(sheet);

    if(worksheet.image_props_.empty() && worksheet.embedded_image_props_.empty() && worksheet.chart_data_.empty() &&
       !worksheet.has_header_vml_ && !worksheet.has_background_image_)
    {
      continue;
    }

    drawing_id++;

    // Prepare embedded worksheet images.
    for(auto& object_props: worksheet.embedded_image_props_)
    {
      store_image_type(object_props.image_type_);

      // Check for images with alt-text.
      if(!object_props.description_.empty())
      {
        has_embedded_image_descriptions_ = true;
      }

      // Check for duplicate images and only store the first instance.
      uint32_t ref_id;
      const auto it = embedded_image_md5_.find(object_props.md5_);
      if(it != std::end(embedded_image_md5_))
      {
        ref_id                     = it->second;
        object_props.is_duplicate_ = true;
      }
      else
      {
        image_ref_id++;
        ref_id = image_ref_id;
        num_embedded_images_++;
        embedded_image_md5_[object_props.md5_] = ref_id;
      }

      worksheet.set_error_cell(object_props, ref_id);
    }

    // Prepare background images.
    if(worksheet.has_background_image_)
    {
      object_properties_t object_props = worksheet.background_image_.value();
      store_image_type(object_props.image_type_);

      uint32_t ref_id;
      const auto it = background_md5_.find(object_props.md5_);
      if(it != std::end(background_md5_))
      {
        ref_id                     = it->second;
        object_props.is_duplicate_ = true;
      }
      else
      {
        image_ref_id++;
        ref_id                             = image_ref_id;
        background_md5_[object_props.md5_] = ref_id;
      }

      worksheet.prepare_background(ref_id, object_props);
    }

    // Prepare worksheet images.
    for(auto& object_props: worksheet.image_props_)
    {
      // Ignore background image added above.
      if(object_props.is_background_)
      {
        continue;
      }

      store_image_type(object_props.image_type_);

      // Check for duplicate images and only store the first instance.
      uint32_t ref_id;
      const auto it = image_md5_.find(object_props.md5_);
      if(it != std::end(image_md5_))
      {
        ref_id                     = it->second;
        object_props.is_duplicate_ = true;
      }
      else
      {
        image_ref_id++;
        ref_id                        = image_ref_id;
        image_md5_[object_props.md5_] = ref_id;
      }

      worksheet.prepare_image(ref_id, drawing_id, object_props);
    }

    // Prepare worksheet charts.
    for(auto& object_props: worksheet.chart_data_)
    {
      chart_ref_id++;
      worksheet.prepare_chart(chart_ref_id, drawing_id, object_props, is_chartsheet);
      if(object_props.chart_)
      {
        ordered_charts_.push_back(object_props.chart_);
      }
    }

    // Prepare worksheet header/footer images.
    for(auto& object_props: worksheet.header_footer_objs_)
    {
      if(object_props)
      {
        store_image_type(object_props->image_type_);

        // Check for duplicate images and only store the first instance.
        uint32_t ref_id;
        const auto it = header_image_md5_.find(object_props->md5_);
        if(it != std::end(header_image_md5_))
        {
          ref_id                      = it->second;
          object_props->is_duplicate_ = true;
        }
        else
        {
          image_ref_id++;
          ref_id                                = image_ref_id;
          header_image_md5_[object_props->md5_] = ref_id;
        }

        worksheet.prepare_header_image(ref_id, *object_props);
      }
    }
  }

  drawing_count_ = drawing_id;
}

void workbook_t::prepare_vml()
{
  uint32_t comment_id     = 0;
  uint32_t vml_drawing_id = 0;
  uint32_t vml_data_id    = 1;
  uint32_t vml_header_id  = 0;
  uint32_t vml_shape_id   = 1024;
  uint32_t comment_count  = 0;

  for(auto& sheet: sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      auto& ws = std::get<worksheet_t>(sheet);

      if(!ws.has_vml_ && !ws.has_header_vml_)
      {
        continue;
      }

      if(ws.has_vml_)
      {
        has_vml_ = true;
        if(ws.has_comments_)
        {
          comment_count_++;
          comment_id++;
          has_comments_ = true;
        }

        vml_drawing_id++;

        comment_count = ws.prepare_vml_objects(vml_data_id, vml_shape_id, vml_drawing_id, comment_id);

        // Each VML should start with a shape id incremented by 1024.
        vml_data_id += 1 * ((1024 + comment_count) / 1024);
        vml_shape_id += 1024 * ((1024 + comment_count) / 1024);
      }

      if(ws.has_header_vml_)
      {
        has_vml_ = true;
        vml_drawing_id++;
        vml_header_id++;
        ws.prepare_header_vml_objects(vml_header_id, vml_drawing_id);
      }
    }
  }
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

  for(auto& sheet: sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      auto& ws = std::get<worksheet_t>(sheet);

      // Check for autofilter settings and store them.
      if(ws.autofilter_.in_use_)
      {
        std::string area = rowcol_to_range_abs(ws.autofilter_.first_row_, ws.autofilter_.first_col_,
                                               ws.autofilter_.last_row_, ws.autofilter_.last_col_);

        // Autofilters are the only defined name to set the hidden flag.
        store_defined_name("_xlnm._FilterDatabase", std::format("{}!_FilterDatabase", ws.quoted_name_),
                           std::format("{}!{}", ws.quoted_name_, area), ws.index_, true);
      }

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
    }
  }
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
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"xWindow",      "240"                         },
      {"yWindow",      "15"                          },
      {"windowWidth",  std::to_string(window_width_) },
      {"windowHeight", std::to_string(window_height_)},
  };

  if(first_sheet_)
  {
    attributes.emplace_back("firstSheet", std::to_string(first_sheet_));
  }

  if(active_sheet_)
  {
    attributes.emplace_back("activeTab", std::to_string(active_sheet_));
  }

  return xml_empty_tag("workbookView", attributes);
}

std::string workbook_t::write_book_views() const
{
  std::string xml_data = xml_start_tag("bookViews");
  xml_data += write_workbook_view();
  xml_data += xml_end_tag("bookViews");

  return xml_data;
}

std::string workbook_t::write_sheet(std::string_view name, uint32_t sheet_id, bool hidden) const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"name", std::string(name)},
      {"sheetId", std::format("{}", sheet_id)},
  };

  if(hidden)
  {
    attributes.emplace_back("state", "hidden");
  }
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
    if(std::holds_alternative<chartsheet_t>(sheet))
    {
      const auto cs = std::get<chartsheet_t>(sheet);
      xml_data += write_sheet(cs.get_sheet_name(), cs.get_sheet_index() + 1, cs.hidden_);
    }
    else if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto ws = std::get<worksheet_t>(sheet);
      xml_data += write_sheet(ws.get_sheet_name(), ws.get_sheet_index() + 1, ws.hidden_);
    }
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

std::string workbook_t::write_defined_name(const defined_name_t& defined_name) const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
      {"name", defined_name.name_}
  };

  if(defined_name.index_ != -1)
  {
    attributes.emplace_back("localSheetId", std::to_string(defined_name.index_));
  }

  if(defined_name.hidden_)
  {
    attributes.emplace_back("hidden", "1");
  }

  return xml_data_element("definedName", defined_name.formula_, attributes);
}

std::string workbook_t::write_defined_names() const
{
  if(defined_names_.empty())
  {
    return "";
  }

  std::string xml_data = xml_start_tag("definedNames");
  for(const auto& defined_name: defined_names_)
  {
    xml_data += write_defined_name(defined_name);
  }
  xml_data += xml_end_tag("definedNames");

  return xml_data;
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

  // Add the default cell format.
  auto format = add_format();
  // Initialize its index.
  get_xf_index(format);

  // Add the default hyperlink format.
  default_url_format_ = add_format();
  default_url_format_->set_hyperlink();
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
      .index_              = num_sheets_,
      .hidden_             = 0,
      ///     .optimize = self->options.constant_memory,
      .active_sheet_       = &active_sheet_,
      .first_sheet_        = &first_sheet_,
      .sst_                = &sst_,
      .name_               = std::string{sheetname},
      .quoted_name_        = quote_sheetname(sheetname),
      ///     .tmpdir = self->options.tmpdir,
      .default_url_format_ = default_url_format_,
      ///     .max_url_length = self->max_url_length,
      ///     .use_1904_epoch = self->use_1904_epoch,
  };

  sheets_.emplace_back(worksheet_t{init_data, std::bind(&workbook_t::get_xf_index, this, std::placeholders::_1),
                                   std::bind(&workbook_t::get_dxf_index, this, std::placeholders::_1)});
  num_worksheets_++;
  num_sheets_++;

  worksheet_names_[init_data.name_] = &std::get<worksheet_t>(sheets_.back());

  return std::get<worksheet_t>(sheets_.back());
}

chartsheet_t& workbook_t::add_chartsheet()
{
  const std::string sheetname = "Chart" + std::to_string(num_chartsheets_ + 1);
  return add_chartsheet(sheetname);
}

chartsheet_t& workbook_t::add_chartsheet(std::string_view sheetname)
{
  // Check that the worksheet name is valid.
  validate_sheetname(sheetname);

  ///     lxw_sheet *sheet = NULL;
  ///     lxw_chartsheet *chartsheet = NULL;
  ///     lxw_chartsheet_name *chartsheet_name = NULL;
  ///     lxw_error error;

  const worksheet_init_data_t init_data{
      .index_        = num_sheets_,
      .hidden_       = 0,
      ///     .optimize = self->options.constant_memory,
      .active_sheet_ = &active_sheet_,
      .first_sheet_  = &first_sheet_,
      .sst_          = &sst_,
      .name_         = std::string{sheetname},
      .quoted_name_  = quote_sheetname(sheetname),
      ///     .tmpdir = self->options.tmpdir,
      /// .default_url_format_ = default_url_format_,
      ///     .max_url_length = self->max_url_length,
      ///     .use_1904_epoch = self->use_1904_epoch,
  };

  sheets_.emplace_back(chartsheet_t{init_data, std::bind(&workbook_t::get_xf_index, this, std::placeholders::_1)});
  num_chartsheets_++;
  num_sheets_++;

  chartsheet_names_[init_data.name_] = &std::get<chartsheet_t>(sheets_.back());

  return std::get<chartsheet_t>(sheets_.back());
}

// TODO No need to create chart through workbook. Can be autonomous object and be added
// to worksheet with insert_chart API (without duplication check)
chart_t& workbook_t::add_chart(chart_type_t chart_type)
{
  if(chart_type == chart_type_t::NONE)
  {
    throw xwpp_exception_t("workbook_t::add_chart: chart type set to 'none'");
  }

  chart_t chart(chart_type);

  charts_.emplace_back(chart);
  return charts_.back();
}

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

  // TODO Add search to be sure there is no duplication.
  format->xf_index_ = static_cast<int32_t>(used_xf_formats_.size());
  used_xf_formats_.push_back(format);

  return format->xf_index_;
}

int32_t workbook_t::get_dxf_index(format_t* format)
{
  // Format already has an index number so return it.
  if(format->dxf_index_ != format_t::PROPERTY_UNSET)
  {
    return format->dxf_index_;
  }

  // TODO Add search to be sure there is no duplication.
  format->dxf_index_ = static_cast<int32_t>(used_dxf_formats_.size());
  used_dxf_formats_.push_back(format);

  return format->dxf_index_;
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
    if(std::holds_alternative<worksheet_t>(sheets_.front()))
    {
      auto& sheet = std::get<worksheet_t>(sheets_.front());
      sheet.select();
    }
  }

  // Set the active sheet and check if a metadata file is needed.
  for(auto& sheet: sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      auto& worksheet = std::get<worksheet_t>(sheet);
      if(worksheet.index_ == active_sheet_)
      {
        worksheet.active_ = true;
      }

      if(worksheet.has_dynamic_functions_)
      {
        has_metadata_          = true;
        has_dynamic_functions_ = true;
      }

      if(!worksheet.embedded_image_props_.empty())
      {
        has_metadata_        = true;
        has_embedded_images_ = true;
      }
    }
  }

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

void workbook_t::define_name(const std::string& name, const std::string& formula)
{
  store_defined_name(name, "", formula, -1, false);
}

void workbook_t::set_properties(const doc_properties_t& properties)
{
  properties_ = properties;
}

void workbook_t::set_custom_property(std::string_view name, const std::string& value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long", name));
  }

  if(value.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): value of custom property '{}' is too long", value));
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
    throw xwpp_out_of_range_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::INTEGER, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, double value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::DOUBLE, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, bool value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::BOOLEAN, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::system_clock::time_point& value)
{
  if(name.empty())
  {
    throw xwpp_out_of_range_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_out_of_range_t(
        std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long", name));
  }

  if(value.time_since_epoch().count() == 0)
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): date is not set");
  }

  custom_properties_.emplace_back(custom_property_types_t::DATETIME, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::year_month_day& value)
{
  set_custom_property(name, std::chrono::sys_days{value} + 0h + 0min + 0s + 0ms);
}

const worksheet_t* workbook_t::get_worksheet_by_name(std::string_view name) const
{
  if(name.empty())
  {
    return nullptr;
  }

  const auto it = worksheet_names_.find(std::string{name});
  if(it != std::end(worksheet_names_))
  {
    return it->second;
  }

  return nullptr;
}

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

format_t* workbook_t::get_default_url_format() const
{
  return default_url_format_;
}

void workbook_t::unset_default_url_format()
{
  default_url_format_->hyperlink_ = false;
  default_url_format_->xf_id_     = 0;
  default_url_format_->underline_ = format_underlines_t::NONE;
  default_url_format_->theme_     = 0;
}

void workbook_t::validate_sheetname(std::string_view sheetname) const
{
  // Check for empty worksheet name.
  if(sheetname.empty())
  {
    throw xwpp_exception_t("workbook_t::validate_sheetname(): sheetname is empty");
  }

  // Check the length of the worksheet name.
  if(sheetname.size() > XWPP_SHEETNAME_MAX)
  {
    throw xwpp_exception_t(std::format("workbook_t::validate_sheetname(): sheetname '{}' is too long", sheetname));
  }

  // Check that the worksheet name doesn't contain invalid characters.
  if(sheetname.find_first_of("[]:*?/\\") != std::string::npos)
  {
    throw xwpp_exception_t(
        std::format("workbook_t::validate_sheetname(): sheetname '{}' contains invalid characters", sheetname));
  }

  // Check that the worksheet doesn't start or end with an apostrophe.
  if(sheetname[0] == '\'' || sheetname[sheetname.size() - 1] == '\'')
  {
    throw xwpp_exception_t(
        std::format("workbook_t::validate_sheetname(): sheetname '{}' contains unbalanced single quote", sheetname));
  }

  // Check if the worksheet name is already in use.
  if(get_worksheet_by_name(sheetname) != nullptr)
  {
    throw xwpp_exception_t(std::format("workbook_t::validate_sheetname(): sheetname '{}' already used", sheetname));
  }

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
