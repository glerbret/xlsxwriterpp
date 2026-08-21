/*
 * rich_value_structure - A library for creating Excel XLSX rich_value_structure files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/rich_value_structure.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_k(const std::string& name, const std::string& type)
{
  return xml_empty_tag("k", {
                              {"n", name},
                              {"t", type},
  });
}

}

rich_value_structure_t::rich_value_structure_t(bool has_embedded_image_descriptions)
  : has_embedded_image_descriptions_{has_embedded_image_descriptions}
{
}

std::string rich_value_structure_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_rv_structures();
  xml_data += xml_end_tag("rvStructures");

  return xml_data;
}

std::string rich_value_structure_t::write_rv_structures() const
{
  std::string xml_data =
    xml_start_tag("rvStructures", {
                                    {"xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"},
                                    {"count", "1"                                                              },
  });
  xml_data += write_s();

  return xml_data;
}

std::string rich_value_structure_t::write_s() const
{
  std::string xml_data = xml_start_tag("s", {
                                              {"t", "_localImage"}
  });
  xml_data += write_k("_rvRel:LocalImageIdentifier", "i");
  xml_data += write_k("CalcOrigin", "i");
  if(has_embedded_image_descriptions_)
  {
    xml_data += write_k("Text", "s");
  }
  xml_data += xml_end_tag("s");

  return xml_data;
}

}