/*
 * rich_value_rel - A library for creating Excel XLSX rich_value_rel files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/rich_value_rel.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>

namespace xwpp
{

rich_value_rel_t::rich_value_rel_t(uint32_t num_embedded_images)
  : num_embedded_images_{num_embedded_images}
{
}

std::string rich_value_rel_t::write_rel(uint32_t rel_index) const
{
  return xml_empty_tag("rel", {
                                  {"r:id", std::format("rId{}", rel_index)}
  });
}

std::string rich_value_rel_t::write_rich_value_rels() const
{
  return xml_start_tag("richValueRels",
                       {
                           {"xmlns",   "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"},
                           {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
  });
}

std::string rich_value_rel_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_rich_value_rels();
  for(uint32_t i = 1; i <= num_embedded_images_; i++)
  {
    xml_data += write_rel(i);
  }
  xml_data += xml_end_tag("richValueRels");

  return xml_data;
}

}
