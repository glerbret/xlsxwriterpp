/*
 * custom - A library for creating Excel custom property files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/custom.h"

#include "xwpp/common.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>
#include <variant>
#include <vector>

namespace xwpp
{

custom_t::custom_t(const std::vector<custom_property_t>& custom_properties)
  : custom_properties_{custom_properties}
{
}

std::string custom_t::assemble_xml_file()
{
  std::string xml_data = xml_declaration();
  xml_data += write_custom_properties();
  xml_data += xml_end_tag("Properties");

  return xml_data;
}

std::string custom_t::write_custom_properties() const
{
  std::string xml_data = xml_start_tag("Properties", {
                                                       {"xmlns",    SCHEMA_OFFICEDOC + "/custom-properties"},
                                                       {"xmlns:vt", SCHEMA_OFFICEDOC + "/docPropsVTypes"   },
  });

  for(size_t index = 1; const auto& property: custom_properties_)
  {
    xml_data += write_custom_property(index, property);
    index++;
  }

  return xml_data;
}

std::string custom_t::write_custom_property(size_t pid, const custom_property_t& property) const
{
  const std::string fmtid = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

  std::string xml_data = xml_start_tag("property", {
                                                     {"fmtid", fmtid                  },
                                                     {"pid",   std::to_string(pid + 1)},
                                                     {"name",  property.name_         },
  });

  if(property.type_ == custom_property_types_t::STRING)
  {
    xml_data += write(std::get<std::string>(property.value_));
  }
  else if(property.type_ == custom_property_types_t::DOUBLE)
  {
    xml_data += write(std::get<double>(property.value_));
  }
  else if(property.type_ == custom_property_types_t::INTEGER)
  {
    xml_data += write(std::get<int32_t>(property.value_));
  }
  else if(property.type_ == custom_property_types_t::BOOLEAN)
  {
    xml_data += write(std::get<bool>(property.value_));
  }
  else if(property.type_ == custom_property_types_t::DATETIME)
  {
    xml_data += write(std::get<std::chrono::system_clock::time_point>(property.value_));
  }

  xml_data += xml_end_tag("property");

  return xml_data;
}

std::string custom_t::write(const std::string& value) const
{
  return xml_data_element("vt:lpwstr", value);
}

std::string custom_t::write(double value) const
{
  return xml_data_element("vt:r8", std::format("{}", value));
}

std::string custom_t::write(int32_t value) const
{
  return xml_data_element("vt:i4", std::to_string(value));
}

std::string custom_t::write(bool value) const
{
  if(value)
  {
    return xml_data_element("vt:bool", "true");
  }
  else
  {
    return xml_data_element("vt:bool", "false");
  }
}

std::string custom_t::write(const std::chrono::system_clock::time_point& value) const
{
  return xml_data_element("vt:filetime",
                          std::format("{:%FT%TZ}", std::chrono::time_point_cast<std::chrono::seconds>(value)));
}

}