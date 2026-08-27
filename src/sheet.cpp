/*
 * sheet - A library for creating Excel XLSX sheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/sheet.h"

#include "xwpp/common.h"
#include "xwpp/exception.h"
#include "xwpp/md5.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <algorithm>
#include <cassert>
#include <filesystem>
#include <format>
#include <fstream>
#include <list>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

using namespace std::literals;

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_drawing(uint16_t id)
{
  return xml_empty_tag("drawing", {
                                    {"r:id", std::format("rId{}", id)}
  });
}

[[nodiscard]] std::string get_vml_image_position(image_position_t image_position)
{
  switch(image_position)
  {
    case image_position_t::HEADER_LEFT:
      return "LH";

    case image_position_t::HEADER_CENTER:
      return "CH";

    case image_position_t::HEADER_RIGHT:
      return "RH";

    case image_position_t::FOOTER_LEFT:
      return "LF";

    case image_position_t::FOOTER_CENTER:
      return "CF";

    case image_position_t::FOOTER_RIGHT:
      return "RF";
  }

  return "";
}

void process_png(object_properties_t& image_props, const std::vector<unsigned char>& data)
{
  uint32_t width  = 0;
  uint32_t height = 0;
  double x_dpi    = 96;
  double y_dpi    = 96;

  // Start after header.
  auto it = std::begin(data);
  it += (4 + 4);

  while(it + static_cast<std::vector<unsigned char>::const_iterator::difference_type>(3 * 4) < std::end(data))
  {
    // Read the PNG length and type fields for the sub-section.
    const uint32_t length = (it[0] << 24U) + (it[1] << 16U) + (it[2] << 8U) + it[3];
    it += 4;

    const std::string type{it, it + 4};
    it += 4;

    if(it + length < std::end(data))
    {
      if(type == "IHDR")
      {
        width = (it[0] << 24U) + (it[1] << 16U) + (it[2] << 8U) + it[3];
        it += 4;

        height = (it[0] << 24U) + (it[1] << 16U) + (it[2] << 8U) + it[3];
        it += 4;

        it += length - (2 * 4);
      }
      else if(type == "pHYs")
      {
        const uint32_t x_ppu = (it[0] << 24U) + (it[1] << 16U) + (it[2] << 8U) + it[3];
        it += 4;

        const uint32_t y_ppu = (it[0] << 24U) + (it[1] << 16U) + (it[2] << 8U) + it[3];
        it += 4;

        const uint8_t units = it[0];
        it++;

        if(units == 1)
        {
          x_dpi = x_ppu * 0.0254;
          y_dpi = y_ppu * 0.0254;
        }

        it += length - ((2 * 4) + 1);
      }
      else if(type == "IEND")
      {
        break;
      }
      else
      {
        it += length;
      }
    }
    else
    {
      // No enough byte for subsection, ==> stop.
      break;
    }

    // Ignore CRC.
    it += 4;
  }

  // Ensure that we read some valid data from the file.
  if(width == 0)
  {
    throw xwpp_exception_t("process_jpeg(): file not valid.");
  }

  // Set the image metadata.
  image_props.image_type_ = image_types_t::PNG;
  image_props.width_      = width;
  image_props.height_     = height;
  image_props.x_dpi_      = (x_dpi != 0.) ? x_dpi : 96;
  image_props.y_dpi_      = (y_dpi != 0.) ? y_dpi : 96;
  image_props.extension_  = "png";
}

void process_jpeg(object_properties_t& image_props, const std::vector<unsigned char>& data)
{
  uint32_t width  = 0;
  uint32_t height = 0;
  double x_dpi    = 96;
  double y_dpi    = 96;

  // Start after header.
  auto it = std::begin(data);
  it += 2;

  // Search through the image data and read the JPEG markers.
  while(it + 4 < std::end(data))
  {
    // Read the JPEG marker and length fields for the sub-section.
    const uint16_t marker = (*it * 0x100) + *(it + 1);
    it += 2;
    const uint16_t length = (*it * 0x100) + *(it + 1);
    it += 2;

    // The offset for next fseek() is the field length + type length.
    uint32_t offset = length - 2;

    // Read the height and width in the 0xFFCn elements (except C4, C8
    // and CC which aren't SOF markers).
    if((marker & 0xFFF0U) == 0xFFC0U && marker != 0xFFC4U && marker != 0xFFC8U && marker != 0xFFCCU)
    {
      if(it + 5 < std::end(data))
      {
        // Skip 1 byte to height and width.
        it++;

        height = (*it * 0x100) + *(it + 1);
        it += 2;
        width = (*it * 0x100) + *(it + 1);
        it += 2;
        offset -= 5;
      }
    }

    // Read the DPI in the 0xFFE0 element.
    if(marker == 0xFFE0)
    {
      if(it + 12 < std::end(data))
      {
        it += 7;
        const uint8_t units = *it;
        it++;
        const uint16_t x_density = (*it * 0x100) + *(it + 1);
        it += 2;
        const uint16_t y_density = (*it * 0x100) + *(it + 1);
        it += 2;

        if(units == 1)
        {
          x_dpi = x_density;
          y_dpi = y_density;
        }

        if(units == 2)
        {
          x_dpi = x_density * 2.54;
          y_dpi = y_density * 2.54;
        }
        offset -= 12;
      }
    }

    if(marker == 0xFFDA)
    {
      break;
    }

    it += offset;
  }

  // Ensure that we read some valid data from the file.
  if(width == 0)
  {
    throw xwpp_exception_t("process_jpeg(): file not valid.");
  }

  // Set the image metadata.
  image_props.image_type_ = image_types_t::JPEG;
  image_props.width_      = width;
  image_props.height_     = height;
  image_props.x_dpi_      = (x_dpi != 0.) ? x_dpi : 96;
  image_props.y_dpi_      = (y_dpi != 0.) ? y_dpi : 96;
  image_props.extension_  = "jpeg";
}

void process_bmp(object_properties_t& image_props, const std::vector<unsigned char>& data)
{
  const double x_dpi = 96;
  const double y_dpi = 96;

  // Skip 18 bytes to the start of the BMP height/width.
  auto it = std::begin(data);
  it += 18;

  const uint32_t width = (*(it + 3) * 0x1000000) + (*(it + 2) * 0x10000) + (*(it + 1) * 0x100) + *it;
  it += 4;

  const auto height =
    static_cast<int32_t>(static_cast<uint32_t>(*(it + 3) * 0x1000000U) + static_cast<uint32_t>(*(it + 2) * 0x10000U) +
                         static_cast<uint32_t>(*(it + 1) * 0x100U) + static_cast<uint32_t>(*it));
  it += 4;

  // Ensure that we read some valid data from the file.
  if(width == 0 || height == 0)
  {
    throw xwpp_exception_t("process_jpeg(): file not valid.");
  }

  // Set the image metadata.
  image_props.image_type_ = image_types_t::BMP;
  image_props.width_      = width;
  // The height can be stored as negative for a top-down DIB so we need to
  // take the absolute value.
  image_props.height_     = height > 0 ? static_cast<uint32_t>(height) : static_cast<uint32_t>(-height);
  image_props.x_dpi_      = x_dpi;
  image_props.y_dpi_      = y_dpi;
  image_props.extension_  = "bmp";
}

void process_gif(object_properties_t& image_props, const std::vector<unsigned char>& data)
{
  const double x_dpi = 96;
  const double y_dpi = 96;

  // Skip 6 bytes to the start of the GIF height/width.
  auto it = std::begin(data);
  it += 6;

  const uint16_t width = (*(it + 1) * 0x100) + *it;
  it += 2;

  const uint16_t height = (*(it + 1) * 0x100) + *it;
  it += 2;

  // Ensure that we read some valid data from the file.
  if(width == 0)
  {
    throw xwpp_exception_t("process_jpeg(): file not valid.");
  }

  // Set the image metadata.
  image_props.image_type_ = image_types_t::GIF;
  image_props.width_      = width;
  image_props.height_     = height;
  image_props.x_dpi_      = x_dpi;
  image_props.y_dpi_      = y_dpi;
  image_props.extension_  = "gif";
}

void process_image(object_properties_t& image_props, const std::vector<unsigned char>& buffer)
{
  if(buffer[1] == 'P' && buffer[2] == 'N' && buffer[3] == 'G')
  {
    process_png(image_props, buffer);
  }
  else if(buffer[0] == 0xFF && buffer[1] == 0xD8)
  {
    process_jpeg(image_props, buffer);
  }
  else if(buffer[0] == 'B' && buffer[1] == 'M')
  {
    process_bmp(image_props, buffer);
  }
  else if(buffer[0] == 'G' && buffer[1] == 'I' && buffer[2] == 'F' && buffer[3] == '8')
  {
    process_gif(image_props, buffer);
  }
  else
  {
    throw xwpp_exception_t(std::format("process_image(): unsupported image format for: {}.", image_props.filename_));
  }

  // Calculate an MD5 checksum for the image so that we can remove duplicate
  // images to reduce the xlsx file size.
  image_props.md5_ = md5_t::digest_to_string(buffer);
}

const uint32_t HEADER_FOOTER_MAX = 255;

}

void sheet_t::set_tab_color(color_t color)
{
  tab_color_ = color;
}

void sheet_t::set_zoom(uint16_t scale)
{
  // Confine the scale to Excel"s range.
  if(scale < 10 || scale > 400)
  {
    throw xwpp_exception_t(
      std::format("sheet_t::set_footer(): zoom factor {} scale outside range: 10 <= zoom <= 400.", scale));
  }

  zoom_ = scale;
}

void sheet_t::set_paper(uint8_t paper_size)
{
  if(paper_size > 118)
  {
    throw xwpp_exception_t(
      std::format("sheet_t::set_paper(): invalid paper size: {}. Valid range is 0-118.", paper_size));
  }

  paper_size_         = paper_size;
  page_setup_changed_ = true;
}

void sheet_t::set_landscape()
{
  orientation_        = drawing_orientation_t::LANDSCAPE;
  page_setup_changed_ = true;
}

void sheet_t::set_portrait()
{
  orientation_        = drawing_orientation_t::PORTRAIT;
  page_setup_changed_ = true;
}

void sheet_t::hide()
{
  hidden_ = true;

  // A hidden ksheet shouldn't be active or selected.
  selected_ = false;

  // If this is active_sheet or first_sheet reset the workbook value.
  if(*first_sheet_ == index_)
  {
    *first_sheet_ = 0;
  }

  if(*active_sheet_ == index_)
  {
    *active_sheet_ = 0;
  }
}

void sheet_t::activate()
{
  selected_ = true;

  // Active sheet can't be hidden.
  hidden_ = false;

  *active_sheet_ = index_;
}

void sheet_t::select()
{
  selected_ = true;

  // Selected sheet can't be hidden.
  hidden_ = false;
}

void sheet_t::set_first_sheet()
{
  // Active sheet can't be hidden.
  hidden_ = false;

  *first_sheet_ = index_;
}

void sheet_t::set_dpi(uint16_t horizontal_dpi, uint16_t vertical_dpi)
{
  if(horizontal_dpi != 0)
  {
    horizontal_dpi_ = horizontal_dpi;
  }
  if(vertical_dpi != 0)
  {
    vertical_dpi_ = vertical_dpi;
  }
}

void sheet_t::set_header(const std::string& str, const std::optional<header_footer_options_t>& options)
{
  if(str.empty())
  {
    throw xwpp_exception_t("sheet_t::set_header(): header must not be empty.");
  }

  if(str.size() > HEADER_FOOTER_MAX)
  {
    throw xwpp_exception_t(
      std::format("sheet_t::set_header(): header '{}' is too long (max: '{}').", str, HEADER_FOOTER_MAX));
  }

  // Count &G placeholders and ensure there are sufficient images.
  uint8_t placeholder_count = 0;
  for(size_t i = 0; i < str.size() - 1; ++i)
  {
    if(str[i] == '&' && str[i + 1] == 'G')
    {
      placeholder_count++;
    }
  }

  if(placeholder_count > 0 && !options)
  {
    throw xwpp_exception_t("sheet_t::set_header(): '&G' placeholders present but no image supplied.");
  }

  if(options)
  {
    uint8_t image_count = 0;

    // Ensure there are enough images to match the placeholders. There is
    // a potential bug where there are sufficient images but in the wrong
    // positions but we don't currently try to deal with that.
    if(!options->image_left_.empty())
    {
      image_count++;
    }
    if(!options->image_center_.empty())
    {
      image_count++;
    }
    if(!options->image_right_.empty())
    {
      image_count++;
    }

    if(placeholder_count != image_count)
    {
      throw xwpp_exception_t(
        "sheet_t::set_header(): number of '&G' placeholders does not match number of supplied images.");
    }

    if(options->margin_ > 0.0)
    {
      margin_header_ = options->margin_;
    }

    set_header_footer_image(options->image_left_, image_position_t::HEADER_LEFT);
    set_header_footer_image(options->image_center_, image_position_t::HEADER_CENTER);
    set_header_footer_image(options->image_right_, image_position_t::HEADER_RIGHT);
  }

  header_                = str;
  header_footer_changed_ = true;
}

void sheet_t::set_header(const std::string& str)
{
  set_header(str, std::nullopt);
}

// TODO Quite similar to set_header. Maybe merged in one generic function
void sheet_t::set_footer(const std::string& str, const std::optional<header_footer_options_t>& options)
{
  if(str.empty())
  {
    throw xwpp_exception_t("sheet_t::set_footer(): footer must not be empty.");
  }

  if(str.size() > HEADER_FOOTER_MAX)
  {
    throw xwpp_exception_t(
      std::format("sheet_t::set_footer(): footer '{}' is too long (max: '{}').", str, HEADER_FOOTER_MAX));
  }

  // Count &G placeholders and ensure there are sufficient images.
  uint8_t placeholder_count = 0;
  for(size_t i = 0; i < str.size() - 1; ++i)
  {
    if(str[i] == '&' && str[i + 1] == 'G')
    {
      placeholder_count++;
    }
  }

  if(placeholder_count > 0 && !options)
  {
    throw xwpp_exception_t("sheet_t::set_footer(): '&G' placeholders present but no image provided.");
  }

  if(options)
  {
    uint8_t image_count = 0;

    // Ensure there are enough images to match the placeholders. There is
    // a potential bug where there are sufficient images but in the wrong
    // positions but we don't currently try to deal with that.
    if(!options->image_left_.empty())
    {
      image_count++;
    }
    if(!options->image_center_.empty())
    {
      image_count++;
    }
    if(!options->image_right_.empty())
    {
      image_count++;
    }

    if(placeholder_count != image_count)
    {
      throw xwpp_exception_t(
        "sheet_t::set_footer(): number of '&G' placeholders does not match number of supplied images.");
    }

    if(options->margin_ > 0.0)
    {
      margin_footer_ = options->margin_;
    }

    set_header_footer_image(options->image_left_, image_position_t::FOOTER_LEFT);
    set_header_footer_image(options->image_center_, image_position_t::FOOTER_CENTER);
    set_header_footer_image(options->image_right_, image_position_t::FOOTER_RIGHT);
  }

  footer_                = str;
  header_footer_changed_ = true;
}

void sheet_t::set_footer(const std::string& str)
{
  set_footer(str, std::nullopt);
}

void sheet_t::set_margins(double left, double right, double top, double bottom)
{
  if(left >= 0)
  {
    margin_left_ = left;
  }

  if(right >= 0)
  {
    margin_right_ = right;
  }

  if(top >= 0)
  {
    margin_top_ = top;
  }

  if(bottom >= 0)
  {
    margin_bottom_ = bottom;
  }
}

sheet_t::sheet_t(bool is_chartsheet, const sheet_init_data_t& init_data)
  : name_{init_data.name_}
  , quoted_name_{init_data.quoted_name_}
  , index_{init_data.index_}
  , is_chartsheet_{is_chartsheet}
  , hidden_{init_data.hidden_}
  , orientation_{is_chartsheet ? drawing_orientation_t::LANDSCAPE : drawing_orientation_t::PORTRAIT}
  , active_sheet_{init_data.active_sheet_}
  , first_sheet_{init_data.first_sheet_}
{
}

void sheet_t::push_chart(const object_properties_t& object_props)
{
  chart_data_.push_back(object_props);
}

void sheet_t::get_image_properties(object_properties_t& image_props)
{
  if(image_props.image_buffer_.empty())
  {
    // Read image.
    std::ifstream image_stream(image_props.filename_, std::ios::binary);
    const std::vector<unsigned char> buffer(std::istreambuf_iterator<char>(image_stream), {});
    process_image(image_props, buffer);
  }
  else
  {
    process_image(image_props, image_props.image_buffer_);
  }
}

uint32_t sheet_t::find_drawing_rel_index(const std::string& target)
{
  if(target.empty())
  {
    return 0;
  }

  auto it = drawing_rel_ids_.find(target);
  if(it != std::end(drawing_rel_ids_))
  {
    return it->second;
  }
  else
  {
    return 0;
  }
}

uint32_t sheet_t::get_drawing_rel_index(const std::string& target)
{
  if(!target.empty())
  {
    const auto it = drawing_rel_ids_.find(target);
    if(it != std::end(drawing_rel_ids_))
    {
      return it->second;
    }
    else
    {
      drawing_rel_id_++;
      drawing_rel_ids_[target] = drawing_rel_id_;
      return drawing_rel_id_;
    }
  }
  else
  {
    drawing_rel_id_++;
    return drawing_rel_id_;
  }
}

uint32_t sheet_t::find_vml_drawing_rel_index(const std::string& target)
{
  if(target.empty())
  {
    return 0;
  }

  auto it = vml_drawing_rel_ids_.find(target);
  if(it != std::end(vml_drawing_rel_ids_))
  {
    return it->second;
  }
  else
  {
    return 0;
  }
}

uint32_t sheet_t::get_vml_drawing_rel_index(const std::string& target)
{
  if(!target.empty())
  {
    const auto it = vml_drawing_rel_ids_.find(target);
    if(it != std::end(vml_drawing_rel_ids_))
    {
      return it->second;
    }
    else
    {
      vml_drawing_rel_id_++;
      vml_drawing_rel_ids_[target] = vml_drawing_rel_id_;
      return vml_drawing_rel_id_;
    }
  }
  else
  {
    vml_drawing_rel_id_++;
    return vml_drawing_rel_id_;
  }
}

void sheet_t::prepare_header_image(uint32_t image_ref_id, object_properties_t& object_props)
{
  add_image_properties(object_props);

  if(find_vml_drawing_rel_index(object_props.md5_) == 0)
  {
    vml_drawing_links_.emplace_back("/image", std::format("../media/image{}.{}", image_ref_id, object_props.extension_),
                                    "");
  }

  vml_obj_t header_image_vml;
  header_image_vml.width_          = static_cast<uint32_t>(object_props.width_);
  header_image_vml.height_         = static_cast<uint32_t>(object_props.height_);
  header_image_vml.x_dpi_          = object_props.x_dpi_;
  header_image_vml.y_dpi_          = object_props.y_dpi_;
  header_image_vml.image_position_ = object_props.image_position_;
  header_image_vml.name_           = object_props.description_;

  // Strip the extension from the filename.
  if(const size_t pos = header_image_vml.name_.find_last_of('.'); pos != std::string::npos)
  {
    header_image_vml.name_ = header_image_vml.name_.substr(0, pos);
  }

  header_image_vml.rel_index_ = get_vml_drawing_rel_index(object_props.md5_);
  header_image_objs_.push_back(header_image_vml);
}

std::string sheet_t::write_tab_color() const
{
  if(tab_color_ == color_t::UNSET)
  {
    return "";
  }

  return xml_empty_tag("tabColor", {
                                     {"rgb", std::format("FF{:06X}", static_cast<uint32_t>(tab_color_) & COLOR_MASK)}
  });
}

std::string sheet_t::write_sheet_views()
{
  std::string xml_data = xml_start_tag("sheetViews");
  xml_data += write_sheet_view();
  xml_data += xml_end_tag("sheetViews");

  return xml_data;
}

std::string sheet_t::write_sheet_protection() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(!protection_.is_configured_)
  {
    return "";
  }

  if(!protection_.hash_.empty())
  {
    attributes.emplace_back("password", protection_.hash_);
  }

  if(!protection_.no_sheet_)
  {
    attributes.emplace_back("sheet", "1");
  }

  if(!protection_.no_content_)
  {
    attributes.emplace_back("content", "1");
  }

  if(!protection_.objects_)
  {
    attributes.emplace_back("objects", "1");
  }

  if(!protection_.scenarios_)
  {
    attributes.emplace_back("scenarios", "1");
  }

  if(protection_.format_cells_)
  {
    attributes.emplace_back("formatCells", "0");
  }

  if(protection_.format_columns_)
  {
    attributes.emplace_back("formatColumns", "0");
  }

  if(protection_.format_rows_)
  {
    attributes.emplace_back("formatRows", "0");
  }

  if(protection_.insert_columns_)
  {
    attributes.emplace_back("insertColumns", "0");
  }

  if(protection_.insert_rows_)
  {
    attributes.emplace_back("insertRows", "0");
  }

  if(protection_.insert_hyperlinks_)
  {
    attributes.emplace_back("insertHyperlinks", "0");
  }

  if(protection_.delete_columns_)
  {
    attributes.emplace_back("deleteColumns", "0");
  }

  if(protection_.delete_rows_)
  {
    attributes.emplace_back("deleteRows", "0");
  }

  if(protection_.no_select_locked_cells_)
  {
    attributes.emplace_back("selectLockedCells", "1");
  }

  if(protection_.sort_)
  {
    attributes.emplace_back("sort", "0");
  }

  if(protection_.autofilter_)
  {
    attributes.emplace_back("autoFilter", "0");
  }

  if(protection_.pivot_tables_)
  {
    attributes.emplace_back("pivotTables", "0");
  }

  if(protection_.no_select_unlocked_cells_)
  {
    attributes.emplace_back("selectUnlockedCells", "1");
  }

  return xml_empty_tag("sheetProtection", attributes);
}

std::string sheet_t::write_page_margins() const
{
  return xml_empty_tag("pageMargins", {
                                        {"left",   std::format("{}", margin_left_)  },
                                        {"right",  std::format("{}", margin_right_) },
                                        {"top",    std::format("{}", margin_top_)   },
                                        {"bottom", std::format("{}", margin_bottom_)},
                                        {"header", std::format("{}", margin_header_)},
                                        {"footer", std::format("{}", margin_footer_)},
  });
}

std::string sheet_t::write_odd_header() const
{
  return xml_data_element("oddHeader", header_);
}

std::string sheet_t::write_odd_footer() const
{
  return xml_data_element("oddFooter", footer_);
}

std::string sheet_t::write_header_footer() const
{
  if(!header_footer_changed_)
  {
    return "";
  }

  std::string xml_data = xml_start_tag("headerFooter");
  if(!header_.empty())
  {
    xml_data += write_odd_header();
  }
  if(!footer_.empty())
  {
    xml_data += write_odd_footer();
  }
  xml_data += xml_end_tag("headerFooter");

  return xml_data;
}

std::string sheet_t::write_drawings()
{
  if(!drawing_)
  {
    return "";
  }

  return write_drawing(rel_count_next_value());
}

// Process a header/footer image and store it in the correct slot.
void sheet_t::set_header_footer_image(const std::string& filename, image_position_t image_position)
{
  // Not all slots will have image files.
  if(filename.empty())
  {
    return;
  }

  // Check that the image file exists and can be opened.
  {
    const std::ifstream image_stream(filename);
    if(!image_stream)
    {
      throw xwpp_exception_t(std::format(
        "sheet_t::set_header_footer_image(): image file '{}' doesn't exist or cannot be opened.", filename));
    }
  }

  // Create a new object to hold the image properties.
  object_properties_t object_props;

  // Copy other options or set defaults.
  object_props.filename_    = filename;
  // Use the filename as the default description, like Excel.
  object_props.description_ = std::filesystem::path(filename).filename().string();

  // Set VML image position string based on the header/footer/position.
  object_props.image_position_ = get_vml_image_position(image_position);

  get_image_properties(object_props);

  // NOLINTNEXTLINE(cppcoreguidelines-pro-bounds-constant-array-index)
  header_footer_objs_[static_cast<size_t>(image_position)] = object_props;
  has_header_vml_                                          = true;
}

void sheet_t::add_drawing_object(const drawing_object_t& drawing_object, uint32_t drawing_id, bool chartsheet)
{
  if(!drawing_)
  {
    drawing_ = drawing_t{};
    if(chartsheet)
    {
      drawing_->embedded_    = false;
      drawing_->orientation_ = orientation_;
    }
    external_drawing_links_.emplace_back("/drawing", std::format("../drawings/drawing{}.xml", drawing_id), "");
  }

  drawing_->add_drawing_object(drawing_object);
}

void sheet_t::add_chart_link(uint32_t chart_ref_id)
{
  drawing_links_.emplace_back("/chart", std::format("../charts/chart{}.xml", chart_ref_id), "");
}

uint32_t sheet_t::add_link_get_index(const std::string& url,
                                     const std::tuple<std::string, std::string, std::string>& relation)
{
  if(find_drawing_rel_index(url) == 0)
  {
    drawing_links_.push_back(relation);
  }

  return get_drawing_rel_index(url);
}

uint32_t sheet_t::add_image_link_get_index(uint32_t image_ref_id, const object_properties_t& object_props)
{
  if(find_drawing_rel_index(object_props.md5_) == 0)
  {
    drawing_links_.emplace_back("/image", std::format("../media/image{}.{}", image_ref_id, object_props.extension_),
                                "");
  }

  return get_drawing_rel_index(object_props.md5_);
}

bool sheet_t::is_outline_changed() const
{
  return false;
}

bool sheet_t::is_selected() const
{
  return selected_;
}

uint16_t sheet_t::get_zoom() const
{
  return zoom_;
}

color_t sheet_t::get_tab_color() const
{
  return tab_color_;
}

uint8_t sheet_t::get_paper_size() const
{
  return paper_size_;
}

bool sheet_t::is_page_setup_changed() const
{
  return page_setup_changed_;
}

void sheet_t::set_page_setup_changed(bool page_setup_changed)
{
  page_setup_changed_ = page_setup_changed;
}

drawing_orientation_t sheet_t::get_orientation() const
{
  return orientation_;
}

uint16_t sheet_t::get_horizontal_dpi() const
{
  return horizontal_dpi_;
}

uint16_t sheet_t::get_vertical_dpi() const
{
  return vertical_dpi_;
}

void sheet_t::set_protection_obj(const protection_obj_t& protection)
{
  protection_ = protection;
}

void sheet_t::add_image_properties(const object_properties_t& image_properties)
{
  image_props_.push_back(image_properties);
}

bool sheet_t::has_header_vml() const
{
  return has_header_vml_;
}

size_t sheet_t::get_table_count() const
{
  return 0;
}

uint16_t sheet_t::rel_count_next_value()
{
  rel_count_++;
  return rel_count_;
}

const std::string& sheet_t::get_sheet_name() const
{
  return name_;
}

size_t sheet_t::get_sheet_index() const
{
  return index_;
}

bool sheet_t::is_chartsheet() const
{
  return is_chartsheet_;
}

bool sheet_t::is_hidden() const
{
  return hidden_;
}

bool sheet_t::has_background_image() const
{
  return false;
}

object_properties_t sheet_t::get_background_image() const
{
  assert(false);
  throw xwpp_exception_t("sheet_t::get_background_image(): cannot have background image.");
}

bool sheet_t::has_embedded_image() const
{
  return false;
}

std::vector<object_properties_t>& sheet_t::get_embedded_image_properties()
{
  assert(false);
  throw xwpp_exception_t("sheet_t::get_background_image(): cannot have embedded image.");
}

const std::vector<object_properties_t>& sheet_t::get_embedded_image_properties() const
{
  assert(false);
  throw xwpp_exception_t("sheet_t::get_background_image(): cannot have embedded image.");
}

const row_num_t sheet_t::ROW_MAX = 1048576;
const col_num_t sheet_t::COL_MAX = 16384;

}
