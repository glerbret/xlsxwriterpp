/*
 * rich_value - A library for creating Excel XLSX rich_value files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/rich_value.h"

#include "xwpp/xmlwriter.h"

namespace xwpp
{

std::string rich_value_t::assemble_xml_file(const workbook_t& workbook) const
{
  std::string xml_data = xml_declaration();
  xml_data += write_rv_data(workbook);
  xml_data += write_images(workbook);
  xml_data += xml_end_tag("rvData");

  return xml_data;
}

std::string rich_value_t::write_rv_data(const workbook_t& workbook) const
{
  return xml_start_tag("rvData", {
                                   {"xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"},
                                   {"count", std::to_string(workbook.num_embedded_images_)                    },
  });
}

std::string rich_value_t::write_images(const workbook_t& workbook) const
{
  uint32_t index = 0;
  uint8_t type   = 5;

  std::string xml_data;
  for(const auto& sheet: workbook.sheets_)
  {
    if(std::holds_alternative<worksheet_t>(sheet))
    {
      const auto& ws = std::get<worksheet_t>(sheet);

      for(const auto& object_props: ws.embedded_image_props_)
      {
        if(!object_props.is_duplicate_)
        {

          if(object_props.decorative_)
          {
            type = 6;
          }

          xml_data += write_rv();
          xml_data += write_v(std::to_string(index));
          xml_data += write_v(std::to_string(type));
          if(!object_props.description_.empty())
          {
            xml_data += write_v(object_props.description_);
          }
          xml_data += xml_end_tag("rv");

          index++;
        }
      }
    }
  }

  return xml_data;
}

std::string rich_value_t::write_rv() const
{
  return xml_start_tag("rv", {
                               {"s", "0"}
  });
}

std::string rich_value_t::write_v(const std::string& value) const
{
  return xml_data_element("v", value);
}

}
