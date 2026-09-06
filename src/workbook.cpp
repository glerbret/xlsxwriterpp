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
#include <fstream>
#include <functional>
#include <ranges>
#include <string>
#include <utility>

using namespace std::literals::chrono_literals;

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_workbook()
{
  return xml_start_tag("workbook", {
                                     {"xmlns",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main"    },
                                     {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/"
                                                 "relationships"},
  });
}

[[nodiscard]] std::string write_sheet(std::string_view name, size_t sheet_id, bool hidden)
{
  attributes_t attributes{
    {
     {"name", std::string(name)},
     {"sheetId", std::format("{}", sheet_id)},
     }
  };

  if(hidden)
  {
    attributes.add_attribute("state", "hidden");
  }
  attributes.add_attribute("r:id", std::format("rId{}", sheet_id));

  return xml_empty_tag("sheet", attributes);
}

[[nodiscard]] std::string write_calc_pr()
{
  return xml_empty_tag("calcPr", {
                                   {"calcId",         "124519"},
                                   {"fullCalcOnLoad", "1"     },
  });
}

}

workbook_t::workbook_t(bool use_zip64)
  : use_zip64_{use_zip64}
{
  // Add the default cell format.
  auto* format = add_format();
  // and initialize its index.
  get_xf_index(format);

  // Add the default hyperlink format.
  // NOLINTNEXTLINE(cppcoreguidelines-prefer-member-initializer)
  default_url_format_ = add_format();
  default_url_format_->set_hyperlink();
}

void workbook_t::save(std::string_view filename)
{
  // Add a default worksheet if none have been added.
  if(sheets_.empty())
  {
    add_worksheet();
  }

  // Ensure that at least one worksheet has been selected.
  if(active_sheet_ == 0)
  {
    auto* sheet = sheets_.front();
    if(!sheet->is_chartsheet())
    {
      dynamic_cast<worksheet_t*>(sheet)->select();
    }
  }

  // Check if a metadata file is needed.
  for(const auto& worksheet: worksheets_)
  {
    if(worksheet.has_dynamic_functions_)
    {
      has_metadata_          = true;
      has_dynamic_functions_ = true;
    }

    if(worksheet.has_embedded_image())
    {
      has_metadata_        = true;
      has_embedded_images_ = true;
    }
  }

  // Set workbook and worksheet VBA codenames if a macro has been added.
  if(!vba_project_.empty())
  {
    if(vba_codename_.empty())
    {
      set_vba_name("ThisWorkbook");
    }

    for(auto& worksheet: worksheets_)
    {
      if(worksheet.get_vba_name().empty())
      {
        worksheet.set_vba_name(std::format("Sheet{}", worksheet.get_sheet_index() + 1));
      }
    }
  }

  // Prepare the element to package.
  prepare_vml();
  prepare_defined_names();
  prepare_drawings();
  add_chart_cache_data();
  prepare_tables();

  // Create a packager object to assemble sub-elements into a zip file.
  packager_t pkg(filename);

  // Create the package.
  pkg.create_package(*this);
}

void workbook_t::set_properties(const doc_properties_t& properties)
{
  properties_ = properties;
}

void workbook_t::set_custom_property(std::string_view name, const std::string& value)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty.");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  if(value.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): value of custom property '{}' is too long.", value));
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
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty.");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::INTEGER, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, double value)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty.");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::DOUBLE, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, bool value)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty.");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  custom_properties_.emplace_back(custom_property_types_t::BOOLEAN, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::system_clock::time_point& value)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  if(value.time_since_epoch().count() == 0)
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): date is not set.");
  }

  custom_properties_.emplace_back(custom_property_types_t::DATETIME, std::string(name), to_datetime(value));
}

void workbook_t::set_custom_property(std::string_view name, const datetime_t& value)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_custom_property(): name of custom property cannot be empty");
  }

  if(name.size() > 255)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::set_custom_property(): name of custom property '{}' is too long.", name));
  }

  datetime_validate(value);

  custom_properties_.emplace_back(custom_property_types_t::DATETIME, std::string(name), value);
}

void workbook_t::set_custom_property(std::string_view name, const std::chrono::year_month_day& value)
{
  set_custom_property(name, std::chrono::sys_days{value} + 0h + 0min + 0s + 0ms);
}

void workbook_t::read_only_recommended()
{
  read_only_ = 2;
}

void workbook_t::use_1904_epoch()
{
  use_1904_epoch_ = true;
}

void workbook_t::set_max_url_length(uint16_t max_url_length)
{
  max_url_length_ = max_url_length;
}

void workbook_t::set_size(uint16_t width, uint16_t height)
{
  // Convert the width/height to twips at 96 dpi.
  if(width != 0)
  {
    window_width_ = width * 1440 / 96;
  }

  if(height != 0)
  {
    window_height_ = height * 1440 / 96;
  }
}

worksheet_t& workbook_t::add_worksheet(std::string_view sheetname)
{
  // Check that the worksheet name is valid.
  validate_sheetname(sheetname);

  const sheet_init_data_t init_data{
    .index_              = sheets_.size(),
    .hidden_             = false,
    .active_sheet_       = &active_sheet_,
    .first_sheet_        = &first_sheet_,
    .sst_                = &sst_,
    .name_               = std::string{sheetname},
    .quoted_name_        = quote_sheetname(sheetname),
    .default_url_format_ = default_url_format_,
    .max_url_length_     = max_url_length_,
    .use_1904_epoch_     = use_1904_epoch_,
  };

  worksheets_.emplace_back(init_data, std::bind_front(&workbook_t::get_xf_index, this),
                           std::bind_front(&workbook_t::get_dxf_index, this));
  worksheet_names_[to_lower(init_data.name_)] = &worksheets_.back();
  sheets_.push_back(&worksheets_.back());

  return worksheets_.back();
}

worksheet_t& workbook_t::add_worksheet()
{
  const std::string sheetname = std::format("Sheet{}", worksheets_.size() + 1);
  return add_worksheet(sheetname);
}

chartsheet_t& workbook_t::add_chartsheet(std::string_view sheetname)
{
  // Check that the worksheet name is valid.
  validate_sheetname(sheetname);

  const sheet_init_data_t init_data{
    .index_              = sheets_.size(),
    .hidden_             = false,
    .active_sheet_       = &active_sheet_,
    .first_sheet_        = &first_sheet_,
    .sst_                = &sst_,
    .name_               = std::string{sheetname},
    .quoted_name_        = quote_sheetname(sheetname),
    .default_url_format_ = default_url_format_,
    .max_url_length_     = max_url_length_,
    .use_1904_epoch_     = use_1904_epoch_,
  };

  chartsheets_.emplace_back(init_data);
  chartsheet_names_[to_lower(init_data.name_)] = &chartsheets_.back();
  sheets_.push_back(&chartsheets_.back());

  return chartsheets_.back();
}

chartsheet_t& workbook_t::add_chartsheet()
{
  const std::string sheetname = std::format("Chart{}", chartsheets_.size() + 1);
  return add_chartsheet(sheetname);
}

const worksheet_t* workbook_t::get_worksheet_by_name(std::string_view name) const
{
  if(name.empty())
  {
    return nullptr;
  }

  if(const auto it = worksheet_names_.find(to_lower(std::string{name})); it != std::end(worksheet_names_))
  {
    return it->second;
  }

  return nullptr;
}

const chartsheet_t* workbook_t::get_chartsheet_by_name(std::string_view name) const
{
  if(name.empty())
  {
    return nullptr;
  }

  if(const auto it = chartsheet_names_.find(to_lower(std::string{name})); it != std::end(chartsheet_names_))
  {
    return it->second;
  }

  return nullptr;
}

void workbook_t::validate_sheetname(std::string_view sheetname) const
{
  // Check for empty worksheet name.
  if(sheetname.empty())
  {
    throw xwpp_exception_t("workbook_t::validate_sheetname(): sheetname is empty.");
  }

  // Check the length of the worksheet name.
  if(sheetname.size() > XWPP_SHEETNAME_MAX)
  {
    throw xwpp_exception_t(std::format("workbook_t::validate_sheetname(): sheetname '{}' is too long.", sheetname));
  }

  // Check that the worksheet name doesn't contain invalid characters.
  if(sheetname.find_first_of("[]:*?/\\") != std::string::npos)
  {
    throw xwpp_exception_t(
      std::format("workbook_t::validate_sheetname(): sheetname '{}' contains invalid characters.", sheetname));
  }

  // Check that the worksheet doesn't start or end with an apostrophe.
  if(sheetname[0] == '\'' || sheetname[sheetname.size() - 1] == '\'')
  {
    throw xwpp_exception_t(
      std::format("workbook_t::validate_sheetname(): sheetname '{}' contains unbalanced single quote.", sheetname));
  }

  // Check if the worksheet name is already in use.
  if(get_worksheet_by_name(sheetname) != nullptr)
  {
    throw xwpp_exception_t(std::format("workbook_t::validate_sheetname(): sheetname '{}' already used.", sheetname));
  }

  // Check if the chartsheet name is already in use.
  if(get_chartsheet_by_name(sheetname) != nullptr)
  {
    throw xwpp_exception_t(std::format("workbook_t::validate_sheetname(): sheetname '{}' already used.", sheetname));
  }
}

// TODO Add class that encapsulate this pointer for interaction with caller. Pointers will only be used inside library.
// TODO Constructor and pointer of this class should be only usable by workbook and worksheet (friendship).
format_t* workbook_t::add_format()
{
  // NOLINTNEXTLINE(modernize-avoid-bind)
  formats_.emplace_back(std::bind_front(&workbook_t::get_dxf_index, this));

  return &formats_.back();
}

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

// TODO No need to create chart through workbook. Can be autonomous object and be added
// to worksheet with insert_chart API (without duplication check).
chart_t& workbook_t::add_chart(chart_type_t chart_type)
{
  if(chart_type == chart_type_t::NONE)
  {
    throw xwpp_exception_t("workbook_t::add_chart: chart type set to 'none'.");
  }

  charts_.emplace_back(chart_type);
  return charts_.back();
}

void workbook_t::add_vba_project(const std::string& filename)
{
  if(filename.empty())
  {
    throw xwpp_exception_t("workbook_t::add_vba_project(): project filename must be specified.");
  }

  // Check that the VBA file exists and can be opened.
  {
    const std::ifstream vba_stream(filename);
    if(!vba_stream)
    {
      throw xwpp_exception_t(
        std::format("workbook_t::add_vba_project(): project file '{}' doesn't exist or cannot be opened.", filename));
    }
  }

  vba_project_ = filename;
}

void workbook_t::add_signed_vba_project(const std::string& vba_project, const std::string& signature)
{
  add_vba_project(vba_project);

  if(signature.empty())
  {
    throw xwpp_exception_t("workbook_t::add_signed_vba_project(): signature filename must be specified.");
  }

  // Check that the vbaProjectSignature file exists and can be opened.
  {
    const std::ifstream signature_file(signature);
    if(!signature_file)
    {
      throw xwpp_exception_t(std::format(
        "workbook_t::add_signed_vba_project(): signature file {} doesn't exist or can't be opened.", signature));
    }
  }

  vba_project_signature_ = signature;
}

void workbook_t::set_vba_name(std::string_view name)
{
  if(name.empty())
  {
    throw xwpp_exception_t("workbook_t::set_vba_name(): 'name' cannot be an empty string.");
  }

  vba_codename_ = name;
}

void workbook_t::define_name(const std::string& name, const std::string& formula)
{
  store_defined_name(name, "", formula, std::numeric_limits<size_t>::max(), false);
}

void workbook_t::set_default_xf_indices()
{
  for(int32_t index{0}; auto& format: formats_)
  {
    // Skip the hyperlink format.
    if(index != 1)
    {
      get_xf_index(&format);
    }

    index++;
  }
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

int32_t workbook_t::get_xf_index(format_t* format)
{
  // Format already has an index number so return it.
  if(format->xf_index_ != format_t::PROPERTY_UNSET)
  {
    return format->xf_index_;
  }

  if(auto it = std::ranges::find_if(
       used_xf_formats_,
       [=](const auto* fmt) {
         return fmt->xf_id_ == format->xf_id_ && fmt->num_format_ == format->num_format_ &&
                fmt->font_name_ == format->font_name_ && fmt->font_scheme_ == format->font_scheme_ &&
                fmt->num_format_index_ == format->num_format_index_ && fmt->font_index_ == format->font_index_ &&
                fmt->has_font_ == format->has_font_ && fmt->has_dxf_font_ == format->has_dxf_font_ &&
                fmt->font_size_ == format->font_size_ && fmt->bold_ == format->bold_ &&
                fmt->italic_ == format->italic_ && fmt->font_color_ == format->font_color_ &&
                fmt->underline_ == format->underline_ && fmt->font_strikeout_ == format->font_strikeout_ &&
                fmt->font_outline_ == format->font_outline_ && fmt->font_shadow_ == format->font_shadow_ &&
                fmt->font_script_ == format->font_script_ && fmt->font_family_ == format->font_family_ &&
                fmt->font_charset_ == format->font_charset_ && fmt->font_condense_ == format->font_condense_ &&
                fmt->font_extend_ == format->font_extend_ && fmt->theme_ == format->theme_ &&
                fmt->hyperlink_ == format->hyperlink_ && fmt->hidden_ == format->hidden_ &&
                fmt->locked_ == format->locked_ && fmt->text_h_align_ == format->text_h_align_ &&
                fmt->text_wrap_ == format->text_wrap_ && fmt->text_v_align_ == format->text_v_align_ &&
                fmt->text_justlast_ == format->text_justlast_ && fmt->rotation_ == format->rotation_ &&
                fmt->fg_color_ == format->fg_color_ && fmt->bg_color_ == format->bg_color_ &&
                fmt->dxf_fg_color_ == format->dxf_fg_color_ && fmt->dxf_bg_color_ == format->dxf_bg_color_ &&
                fmt->pattern_ == format->pattern_ && fmt->has_fill_ == format->has_fill_ &&
                fmt->has_dxf_fill_ == format->has_dxf_fill_ && fmt->fill_index_ == format->fill_index_ &&
                fmt->fill_count_ == format->fill_count_ && fmt->border_index_ == format->border_index_ &&
                fmt->has_border_ == format->has_border_ && fmt->has_dxf_border_ == format->has_dxf_border_ &&
                fmt->border_count_ == format->border_count_ && fmt->bottom_ == format->bottom_ &&
                fmt->diag_border_ == format->diag_border_ && fmt->diag_type_ == format->diag_type_ &&
                fmt->left_ == format->left_ && fmt->right_ == format->right_ && fmt->top_ == format->top_ &&
                fmt->bottom_color_ == format->bottom_color_ && fmt->diag_color_ == format->diag_color_ &&
                fmt->left_color_ == format->left_color_ && fmt->right_color_ == format->right_color_ &&
                fmt->top_color_ == format->top_color_ && fmt->indent_ == format->indent_ &&
                fmt->shrink_ == format->shrink_ && fmt->merge_range_ == format->merge_range_ &&
                fmt->reading_order_ == format->reading_order_ && fmt->just_distrib_ == format->just_distrib_ &&
                fmt->color_indexed_ == format->color_indexed_ && fmt->font_only_ == format->font_only_ &&
                fmt->quote_prefix_ == format->quote_prefix_;
       });
     it != std::end(used_xf_formats_))
  {
    return (*it)->xf_index_;
  }

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

void workbook_t::prepare_fonts()
{
  std::vector<format_t*> fonts;

  // TODO Use unordered_set to optimise this search.
  for(auto* format: used_xf_formats_)
  {
    for(const auto* font: fonts)
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
  for(auto* format: used_dxf_formats_)
  {
    // The only font properties that can change for a DXF format are:
    // color, bold, italic, underline and strikethrough.
    if(format->font_color_ != color_t::UNSET || format->bold_ || format->italic_ ||
       format->underline_ != format_underlines_t::NONE || format->font_strikeout_)
    {
      format->has_dxf_font_ = true;
    }
  }

  font_count_ = static_cast<uint16_t>(fonts.size());
}

void workbook_t::prepare_num_formats()
{
  std::vector<format_t*> num_formats;

  // TODO Use unordered_set to optimise this search.
  for(auto* format: used_xf_formats_)
  {
    // Format already has a number format index.
    if(format->num_format_index_ != 0)
    {
      continue;
    }

    // Check if there is a user defined number format string.
    if(!format->num_format_.empty())
    {
      for(const auto* num_format: num_formats)
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
        format->num_format_index_ = static_cast<uint16_t>(num_formats.size()) + 0xA4;
        num_formats.push_back(format);
      }
    }
  }

  // Don't update num_format_count for DXF formats.
  num_format_count_ = static_cast<uint16_t>(num_formats.size());

  for(auto* format: used_dxf_formats_)
  {
    // Format already has a number format index.
    if(format->num_format_index_ != 0)
    {
      continue;
    }

    // Check if there is a user defined number format string.
    if(!format->num_format_.empty())
    {
      for(const auto* num_format: num_formats)
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
        format->num_format_index_ = static_cast<uint16_t>(num_formats.size()) + 0xA4;
        num_formats.push_back(format);
      }
    }
  }
}

void workbook_t::prepare_borders()
{
  std::vector<format_t*> borders;

  // TODO Use unordered_set to optimise this search.
  for(auto* format: used_xf_formats_)
  {
    for(const auto* border: borders)
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

  // For DXF formats we only need to check if the properties have changed.
  for(auto* format: used_dxf_formats_)
  {
    if(format->left_ != format_borders_t::NONE || format->right_ != format_borders_t::NONE ||
       format->top_ != format_borders_t::NONE || format->bottom_ != format_borders_t::NONE)
    {
      format->has_dxf_border_ = true;
    }
  }

  border_count_ = static_cast<uint16_t>(borders.size());
}

void workbook_t::prepare_vml()
{
  uint32_t comment_id{0};
  uint32_t vml_drawing_id{0};
  uint32_t vml_data_id{1};
  uint32_t vml_header_id{0};
  uint32_t vml_shape_id{1024};
  uint32_t comment_count{0};

  for(auto& worksheet: worksheets_)
  {
    if(!worksheet.has_vml_ && !worksheet.has_header_vml())
    {
      continue;
    }

    if(worksheet.has_vml_)
    {
      has_vml_ = true;
      if(worksheet.has_comments_)
      {
        comment_count_++;
        comment_id++;
        has_comments_ = true;
      }

      vml_drawing_id++;

      comment_count = worksheet.prepare_vml_objects(vml_data_id, vml_shape_id, vml_drawing_id, comment_id);

      // Each VML should start with a shape id incremented by 1024.
      vml_data_id += 1 * ((1024 + comment_count) / 1024);
      vml_shape_id += 1024 * ((1024 + comment_count) / 1024);
    }

    if(worksheet.has_header_vml())
    {
      has_vml_ = true;
      vml_drawing_id++;
      vml_header_id++;
      worksheet.prepare_header_vml_objects(vml_header_id, vml_drawing_id);
    }
  }
}

/*
 * Process and store the defined names. The defined names are stored with
 * the Workbook.xml but also with the App.xml if they refer to a sheet
 * range like "Sheet1!:A1". The defined names are store in sorted
 * order for consistency with Excel. The names need to be normalized before
 * sorting.
 */
void workbook_t::store_defined_name(std::string_view name, std::string_view app_name, std::string_view formula,
                                    size_t index, bool hidden)
{
  // Do some checks on the input data.
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
    auto worksheet_name = std::string{name.substr(0, found_string)};
    const auto tmp_str  = name.substr(found_string + 1);

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
    for(const auto& worksheet: worksheets_)
    {
      if(worksheet_name == worksheet.name_)
      {
        defined_name.index_                = static_cast<int16_t>(worksheet.index_);
        defined_name.normalised_sheetname_ = worksheet_name;
      }
    }

    // If we didn't find the worksheet name we exit.
    if(defined_name.index_ == std::numeric_limits<size_t>::max())
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

  std::ranges::transform(defined_name.normalised_name_, std::begin(defined_name.normalised_name_),
                         [](char c) { return std::tolower(c); });
  std::ranges::transform(defined_name.normalised_sheetname_, std::begin(defined_name.normalised_sheetname_),
                         [](char c) { return std::tolower(c); });

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

  // Insert at the end otherwise.
  defined_names_.push_back(defined_name);
}

void workbook_t::prepare_defined_names()
{
  for(auto& worksheet: worksheets_)
  {
    // Check for autofilter settings and store them.
    if(worksheet.autofilter_.in_use_)
    {
      std::string area = rowcol_to_range_abs(worksheet.autofilter_.first_row_, worksheet.autofilter_.first_col_,
                                             worksheet.autofilter_.last_row_, worksheet.autofilter_.last_col_);

      // Autofilters are the only defined name to set the hidden flag.
      store_defined_name("_xlnm._FilterDatabase", std::format("{}!_FilterDatabase", worksheet.quoted_name_),
                         std::format("{}!{}", worksheet.quoted_name_, area), worksheet.index_, true);
    }

    // Check for Print Area settings and store them.
    if(worksheet.print_area_.in_use_)
    {
      std::string area;

      // Check for print area that is the max row range.
      if(worksheet.print_area_.first_row_ == 0 && worksheet.print_area_.last_row_ == worksheet_t::ROW_MAX - 1)
      {
        const std::string first_col = col_to_name(worksheet.print_area_.first_col_, false);
        const std::string last_col  = col_to_name(worksheet.print_area_.last_col_, false);
        area                        = std::format("${}:${}", first_col, last_col);
      }
      // Check for print area that is the max column range.
      else if(worksheet.print_area_.first_col_ == 0 &&
              std::cmp_equal(worksheet.print_area_.last_col_, worksheet_t::COL_MAX - 1))
      {
        area = std::format("${}:${}", worksheet.print_area_.first_row_ + 1, worksheet.print_area_.last_row_ + 1);
      }
      else
      {
        area = rowcol_to_range_abs(worksheet.print_area_.first_row_, worksheet.print_area_.first_col_,
                                   worksheet.print_area_.last_row_, worksheet.print_area_.last_col_);
      }

      const std::string app_name = std::format("{}!Print_Area", worksheet.quoted_name_);
      const std::string range    = std::format("{}!{}", worksheet.quoted_name_, area);
      store_defined_name("_xlnm.Print_Area", app_name, range, static_cast<int16_t>(worksheet.index_), false);
    }

    // Check for repeat rows/cols. aka, Print Titles and store them.
    if(worksheet.repeat_rows_.in_use_ || worksheet.repeat_cols_.in_use_)
    {
      if(worksheet.repeat_rows_.in_use_ && worksheet.repeat_cols_.in_use_)
      {
        const std::string app_name  = std::format("{}!Print_Titles", worksheet.quoted_name_);
        const std::string first_col = col_to_name(worksheet.repeat_cols_.first_col_, false);
        const std::string last_col  = col_to_name(worksheet.repeat_cols_.last_col_, false);
        const std::string range =
          std::format("{}!${}:${},{}!${}:${}", worksheet.quoted_name_, first_col, last_col, worksheet.quoted_name_,
                      worksheet.repeat_rows_.first_row_ + 1, worksheet.repeat_rows_.last_row_ + 1);
        store_defined_name("_xlnm.Print_Titles", app_name, range, worksheet.index_, false);
      }
      else if(worksheet.repeat_rows_.in_use_)
      {
        const std::string app_name = std::format("{}!Print_Titles", worksheet.quoted_name_);
        const std::string range =
          std::format("{}!${}:${}", worksheet.quoted_name_, worksheet.repeat_rows_.first_row_ + 1,
                      worksheet.repeat_rows_.last_row_ + 1);
        store_defined_name("_xlnm.Print_Titles", app_name, range, worksheet.index_, false);
      }
      else if(worksheet.repeat_cols_.in_use_)
      {
        const std::string app_name  = std::format("{}!Print_Titles", worksheet.quoted_name_);
        const std::string first_col = col_to_name(worksheet.repeat_cols_.first_col_, false);
        const std::string last_col  = col_to_name(worksheet.repeat_cols_.last_col_, false);
        const std::string range     = std::format("{}!${}:${}", worksheet.quoted_name_, first_col, last_col);
        store_defined_name("_xlnm.Print_Titles", app_name, range, worksheet.index_, false);
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
  uint32_t chart_ref_id{0};
  uint32_t image_ref_id{0};
  uint32_t drawing_id{0};

  for(auto* sheet: sheets_)
  {
    if(sheet->image_props_.empty() && !sheet->has_embedded_image() && sheet->chart_data_.empty() &&
       !sheet->has_header_vml() && !sheet->has_background_image())
    {
      continue;
    }

    drawing_id++;

    // Prepare embedded worksheet images.
    if(sheet->has_embedded_image())
    {
      for(auto& object_props: sheet->get_embedded_image_properties())
      {
        store_image_type(object_props.image_type_);

        // Check for images with alt-text.
        if(!object_props.description_.empty())
        {
          has_embedded_image_descriptions_ = true;
        }

        // Check for duplicate images and only store the first instance.
        uint32_t ref_id{0};
        if(const auto it = embedded_image_md5_.find(object_props.md5_); it != std::end(embedded_image_md5_))
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

        sheet->set_error_cell(object_props, ref_id);
      }
    }

    // Prepare background images.
    if(sheet->has_background_image())
    {
      auto object_props = sheet->get_background_image();
      store_image_type(object_props.image_type_);

      uint32_t ref_id{0};
      if(const auto it = background_md5_.find(object_props.md5_); it != std::end(background_md5_))
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

      sheet->prepare_background(ref_id, object_props);
    }

    // Prepare worksheet images.
    for(auto& object_props: sheet->image_props_)
    {
      // Ignore background image added above.
      if(object_props.is_background_)
      {
        continue;
      }

      store_image_type(object_props.image_type_);

      // Check for duplicate images and only store the first instance.
      uint32_t ref_id{0};
      if(const auto it = image_md5_.find(object_props.md5_); it != std::end(image_md5_))
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

      sheet->prepare_image(ref_id, drawing_id, object_props);
    }

    // Prepare worksheet charts.
    for(auto& object_props: sheet->chart_data_)
    {
      chart_ref_id++;
      sheet->prepare_chart(chart_ref_id, drawing_id, object_props);
      if(object_props.chart_)
      {
        ordered_charts_.push_back(object_props.chart_);
      }
    }

    // Prepare worksheet header/footer images.
    for(auto& object_props: sheet->header_footer_objs_)
    {
      if(object_props)
      {
        store_image_type(object_props->image_type_);

        // Check for duplicate images and only store the first instance.
        uint32_t ref_id{0};
        if(const auto it = header_image_md5_.find(object_props->md5_); it != std::end(header_image_md5_))
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

        sheet->prepare_header_image(ref_id, *object_props);
      }
    }
  }

  drawing_count_ = static_cast<uint16_t>(drawing_id);
}

/* Convert a chart range such as Sheet1!$A$1:$A$5 to a sheet name and row-col
 * dimensions, or vice-versa. This gives us the dimensions to read data back
 * from the worksheet.
 */
void workbook_t::populate_range_dimensions(series_range_t& range) const
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

  // Ignore non-contiguous range like (Sheet1!$A$1:$A$2,Sheet1!$A$4:$A$5).
  if(range.formula_[0] == '(')
  {
    range.ignore_cache_ = true;
    return;
  }

  // Create a copy of the formula to modify and parse into parts.
  const std::string formula = range.formula_;

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
    std::string sheetname     = formula.substr(0, found_string);
    const std::string tmp_str = formula.substr(found_string + 1);

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
        std::format("workbook_t::populate_range_dimensions(): sheetname '{}' in chart formula '{}' doesn't exist.",
                    sheetname, range.formula_));
    }

    range.sheetname_ = sheetname;
    range.first_row_ = name_to_row(tmp_str);
    range.first_col_ = name_to_col(tmp_str);

    found_string = formula.find(':');
    if(found_string == std::string::npos)
    {
      // 1D range.
      range.last_row_ = range.first_row_;
      range.last_col_ = range.first_col_;
    }
    else
    {
      // 2D range.
      range.last_row_ = name_to_row_2(tmp_str);
      range.last_col_ = name_to_col_2(tmp_str);
    }
  }
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
void workbook_t::populate_range_data_cache(series_range_t& range) const
{
  uint16_t num_data_points{0};

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
      std::format("workbook_t::populate_range_data_cache(): sheetname '{}' in chart formula '{}' doesn't exist.",
                  range.sheetname_, range.formula_));
  }

  // Iterate through the worksheet data and populate the range cache.
  for(row_num_t row_num = range.first_row_; row_num <= range.last_row_; row_num++)
  {
    const row_t* row_obj = worksheet->find_row(row_num);
    for(col_num_t col_num = range.first_col_; col_num <= range.last_col_; col_num++)
    {
      series_data_point_t data_point;
      if(const cell_t* cell_obj = worksheet->find_cell_in_row(row_obj, col_num); cell_obj)
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

void workbook_t::populate_range(series_range_t& range) const
{
  populate_range_dimensions(range);
  populate_range_data_cache(range);
}

void workbook_t::add_chart_cache_data() const
{
  for(auto* chart: ordered_charts_)
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

void workbook_t::prepare_tables()
{
  for(uint32_t table_id{0}; auto& worksheet: worksheets_)
  {
    if(!worksheet.table_objs_.empty())
    {
      worksheet.prepare_tables(table_id + 1);
      table_id += static_cast<uint32_t>(worksheet.table_objs_.size());
    }
  }
}

void workbook_t::prepare_fills()
{
  std::vector<format_t*> fills;

  // Add the default fills.
  // NOLINTNEXTLINE(modernize-avoid-bind)
  format_t default_fill_1(std::bind_front(&workbook_t::get_dxf_index, this));
  default_fill_1.fg_color_   = color_t::UNSET;
  default_fill_1.bg_color_   = color_t::UNSET;
  default_fill_1.pattern_    = format_patterns_t::NONE;
  default_fill_1.fill_index_ = 0;
  fills.push_back(&default_fill_1);
  // NOLINTNEXTLINE(modernize-avoid-bind)
  format_t default_fill_2(std::bind_front(&workbook_t::get_dxf_index, this));
  default_fill_2.fg_color_   = color_t::UNSET;
  default_fill_2.bg_color_   = color_t::UNSET;
  default_fill_2.pattern_    = format_patterns_t::GRAY_125;
  default_fill_2.fill_index_ = 1;
  fills.push_back(&default_fill_2);

  // For DXF formats we only need to check if the properties have changed.
  for(auto* format: used_dxf_formats_)
  {
    if(format->pattern_ != format_patterns_t::NONE || format->bg_color_ != color_t::UNSET ||
       format->fg_color_ != color_t::UNSET)
    {
      format->has_dxf_fill_ = true;
      format->dxf_bg_color_ = format->bg_color_;
      format->dxf_fg_color_ = format->fg_color_;
    }
  }

  // TODO Use unordered_set to optimise this search.
  for(auto* format: used_xf_formats_)
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

    for(const auto* fill: fills)
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

  fill_count_ = static_cast<uint16_t>(fills.size());
}

void workbook_t::prepare_workbook()
{
  prepare_fonts();
  prepare_num_formats();
  prepare_borders();
  prepare_fills();
}

std::string workbook_t::write_file_version() const
{
  attributes_t attributes{
    {
     {"appName", "xl"},
     {"lastEdited", "4"},
     {"lowestEdited", "4"},
     {"rupBuild", "4505"},
     }
  };

  if(!vba_project_.empty())
  {
    attributes.add_attribute("codeName", "{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}");
  }

  return xml_empty_tag("fileVersion", attributes);
}

std::string workbook_t::write_file_sharing() const
{
  if(read_only_ == 0)
  {
    return "";
  }

  return xml_empty_tag("fileSharing", {
                                        {"readOnlyRecommended", "1"}
  });
}

std::string workbook_t::write_workbook_pr() const
{
  attributes_t attributes;

  if(!vba_codename_.empty())
  {
    attributes.add_attribute("codeName", vba_codename_);
  }

  if(use_1904_epoch_)
  {
    attributes.add_attribute("date1904", "1");
  }

  attributes.add_attribute("defaultThemeVersion", "124226");

  return xml_empty_tag("workbookPr", attributes);
}

std::string workbook_t::write_workbook_view() const
{
  attributes_t attributes{
    {
     {"xWindow", "240"},
     {"yWindow", "15"},
     {"windowWidth", std::to_string(window_width_)},
     {"windowHeight", std::to_string(window_height_)},
     }
  };

  if(first_sheet_ != 0)
  {
    attributes.add_attribute("firstSheet", first_sheet_);
  }

  if(active_sheet_ != 0)
  {
    attributes.add_attribute("activeTab", active_sheet_);
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

std::string workbook_t::write_sheets() const
{
  std::string xml_data = xml_start_tag("sheets");
  for(const auto* sheet: sheets_)
  {
    xml_data += write_sheet(sheet->get_sheet_name(), sheet->get_sheet_index() + 1, sheet->is_hidden());
  }
  xml_data += xml_end_tag("sheets");

  return xml_data;
}

std::string workbook_t::write_defined_name(const defined_name_t& defined_name)
{
  attributes_t attributes{{{"name", defined_name.name_}}};

  if(defined_name.index_ != std::numeric_limits<size_t>::max())
  {
    attributes.add_attribute("localSheetId", defined_name.index_);
  }

  if(defined_name.hidden_)
  {
    attributes.add_attribute("hidden", "1");
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

}
