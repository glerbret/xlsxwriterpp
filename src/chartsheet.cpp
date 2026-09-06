/*
 * chartsheet - A library for creating Excel XLSX chartsheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/chartsheet.h"

#include "xwpp/exception.h"
#include "xwpp/utility.h"
#include "xwpp/worksheet.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>

namespace xwpp
{

chartsheet_t::chartsheet_t(const sheet_init_data_t& init_data)
  : sheet_t{true, init_data}
{
}

void chartsheet_t::set_chart(chart_t* chart, const std::optional<chart_options_t>& options)
{
  if(!chart)
  {
    throw xwpp_exception_t("chartsheet_t::set_chart(): 'chart' must be present");
  }

  if(chart->in_use_)
  {
    throw xwpp_exception_t(
      "chartsheet_t::set_chart(): the same chart object cannot be inserted in a worksheet more than once");
  }

  if(chart->series_list_.empty())
  {
    throw xwpp_exception_t("chartsheet_t::set_chart(): chart must have a series");
  }

  object_properties_t object_props;
  if(options.has_value())
  {
    object_props.x_offset_ = options->x_offset_;
    object_props.y_offset_ = options->y_offset_;
    object_props.x_scale_  = options->x_scale_;
    object_props.y_scale_  = options->y_scale_;
  }

  object_props.width_  = 480;
  object_props.height_ = 288;

  if(object_props.x_scale_ == 0.0)
  {
    object_props.x_scale_ = 1;
  }

  if(object_props.y_scale_ == 0.0)
  {
    object_props.y_scale_ = 1;
  }

  // Store chart references so they can be ordered in the workbook.
  object_props.chart_ = chart;

  push_chart(object_props);

  chart->in_use_        = true;
  chart->is_chartsheet_ = true;
  chart->is_protected_  = is_protected_;
  chart_                = chart;
}

void chartsheet_t::set_chart(chart_t* chart)
{
  set_chart(chart, std::nullopt);
}

void chartsheet_t::protect(const std::string& password, std::optional<protection_t> options)
{
  // Copy any user parameters to the internal structure.
  protection_obj_t protection;

  if(options.has_value())
  {
    protection.objects_    = options->no_objects_;
    protection.no_content_ = options->no_content_;
  }
  else
  {
    protection.objects_    = false;
    protection.no_content_ = false;
  }

  if(!password.empty())
  {
    protection.hash_ = std::format("{:04X}", hash_password(password));
  }
  else
  {
    if(protection.objects_ && protection.no_content_)
    {
      return;
    }
  }

  protection.no_sheet_      = true;
  protection.scenarios_     = true;
  protection.is_configured_ = true;

  if(chart_)
  {
    chart_->is_protected_ = true;
  }
  else
  {
    is_protected_ = true;
  }

  set_protection_obj(protection);
}

void chartsheet_t::protect(const std::string& password)
{
  protect(password, std::nullopt);
}

void chartsheet_t::protect(std::optional<protection_t> options)
{
  protect("", options);
}

void chartsheet_t::protect()
{
  protect("", std::nullopt);
}

[[nodiscard]] std::string chartsheet_t::assemble_xml_file()
{
  std::string xml_data = xml_declaration();
  xml_data += write_chartsheet();
  xml_data += write_sheet_pr();
  xml_data += write_sheet_views();
  xml_data += write_sheet_protection();
  xml_data += write_page_margins();
  xml_data += write_page_setup();
  xml_data += write_header_footer();
  xml_data += write_drawings();
  xml_data += xml_end_tag("chartsheet");

  return xml_data;
}

std::string chartsheet_t::write_chartsheet()
{
  return xml_start_tag("chartsheet",
                       {
                         {"xmlns",   "http://schemas.openxmlformats.org/spreadsheetml/2006/main"          },
                         {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
  });
}

std::string chartsheet_t::write_sheet_view()
{
  attributes_t attributes;

  // Show that the sheet tab is selected.
  if(is_selected())
  {
    attributes.add_attribute("tabSelected", "1");
  }

  // Set the zoom level.
  if(get_zoom() != 100)
  {
    attributes.add_attribute("zoomScale", get_zoom());
  }

  attributes.add_attribute("workbookViewId", "0");

  return xml_empty_tag("sheetView", attributes);
}

std::string chartsheet_t::write_page_setup() const
{
  if(!is_page_setup_changed())
  {
    return "";
  }

  attributes_t attributes;

  // Set paper size.
  if(get_paper_size() != 0)
  {
    attributes.add_attribute("paperSize", get_paper_size());
  }

  // Set page orientation.
  if(get_orientation() == drawing_orientation_t::PORTRAIT)
  {
    attributes.add_attribute("orientation", "portrait");
  }
  else
  {
    attributes.add_attribute("orientation", "landscape");
  }

  // Set the DPI. Mainly only for testing.
  if(get_horizontal_dpi() != 0)
  {
    attributes.add_attribute("horizontalDpi", get_horizontal_dpi());
  }

  if(get_vertical_dpi() != 0)
  {
    attributes.add_attribute("verticalDpi", get_vertical_dpi());
  }

  return xml_empty_tag("pageSetup", attributes);
}

std::string chartsheet_t::write_sheet_pr() const
{
  if(get_tab_color() != color_t::UNSET || is_outline_changed())
  {
    std::string xml_data = xml_start_tag("sheetPr");
    xml_data += write_tab_color();
    xml_data += xml_end_tag("sheetPr");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("sheetPr");
  }
}

void chartsheet_t::set_error_cell([[maybe_unused]] const object_properties_t& object_props,
                                  [[maybe_unused]] uint32_t ref_id)
{
  throw xwpp_exception_t("chartsheet_t::set_error_cell(): no cell in chartsheet");
}

void chartsheet_t::prepare_background([[maybe_unused]] uint32_t image_ref_id,
                                      [[maybe_unused]] object_properties_t& object_props)
{
  throw xwpp_exception_t("chartsheet_t::prepare_background(): background image not supported in chartsheet");
}

void chartsheet_t::prepare_image(uint32_t image_ref_id, uint32_t drawing_id, object_properties_t& object_props)
{
  drawing_object_t drawing_object{
    .type_          = drawing_types_t::IMAGE,
    .anchor_        = static_cast<uint8_t>(object_position_t::MOVE_DONT_SIZE),
    .rel_index_     = 0,
    .url_rel_index_ = 0,
    .description_   = object_props.description_,
    .tip_           = object_props.tip_,
    .decorative_    = object_props.decorative_,
  };

  // Scale to user scale.
  double width{object_props.width_ * object_props.x_scale_};
  double height{object_props.height_ * object_props.y_scale_};

  // Scale by non 96dpi resolutions.
  width *= 96.0 / object_props.x_dpi_;
  height *= 96.0 / object_props.y_dpi_;

  object_props.width_  = width;
  object_props.height_ = height;

  // Convert from pixels to emus.
  drawing_object.width_  = static_cast<uint32_t>(0.5 + (width * 9525));
  drawing_object.height_ = static_cast<uint32_t>(0.5 + (height * 9525));

  drawing_object.rel_index_ = add_image_link_get_index(image_ref_id, object_props);
  add_drawing_object(drawing_object, drawing_id, true);
}

void chartsheet_t::prepare_chart(uint32_t chart_ref_id, uint32_t drawing_id, object_properties_t& object_props)
{
  drawing_object_t drawing_object{
    .type_          = drawing_types_t::CHART,
    .anchor_        = static_cast<uint8_t>(object_position_t::MOVE_AND_SIZE),
    .rel_index_     = get_drawing_rel_index(""),
    .url_rel_index_ = 0,
    .description_   = object_props.description_,
    .tip_           = "",
    .decorative_    = object_props.decorative_,
  };

  // Scale to user scale.
  const double width{object_props.width_ * object_props.x_scale_};
  const double height{object_props.height_ * object_props.y_scale_};

  // Convert to the nearest pixel.
  object_props.width_  = width;
  object_props.height_ = height;

  // Convert from pixels to emus.
  drawing_object.width_  = static_cast<uint32_t>(0.5 + (width * 9525));
  drawing_object.height_ = static_cast<uint32_t>(0.5 + (height * 9525));

  add_drawing_object(drawing_object, drawing_id, true);
  add_chart_link(chart_ref_id);
}

}
