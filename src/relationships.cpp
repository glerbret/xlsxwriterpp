/*
 * relationships - A library for creating Excel XLSX relationships files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/relationships.h"

#include "xwpp/common.h"
#include "xwpp/xmlwriter.h"

#include <format>

namespace xwpp
{

void relationships_t::add_document(std::string_view type, std::string_view target)
{
  add(SCHEMA_DOCUMENT, type, target);
}

void relationships_t::add_package(std::string_view type, std::string_view target)
{
  add(SCHEMA_PACKAGE, type, target);
}

void relationships_t::add_worksheet(std::string_view type, std::string_view target, std::string_view target_mode)
{
  add(SCHEMA_DOCUMENT, type, target, target_mode);
}

void relationships_t::add_rich_value()
{
  add("http://schemas.microsoft.com/office/2022/10/relationships/", "richValueRel", "richData/richValueRel.xml");
  add("http://schemas.microsoft.com/office/2017/06/relationships/", "rdRichValue", "richData/rdrichvalue.xml");
  add("http://schemas.microsoft.com/office/2017/06/relationships/", "rdRichValueStructure",
      "richData/rdrichvaluestructure.xml");
  add("http://schemas.microsoft.com/office/2017/06/relationships/", "rdRichValueTypes",
      "richData/rdRichValueTypes.xml");
}

void relationships_t::add_ms_package(std::string_view type, std::string_view target)
{
  add(SCHEMA_MS, type, target);
}

std::string relationships_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += xml_start_tag("Relationships", {
                                               {"xmlns", SCHEMA_PACKAGE}
  });

  for(size_t rel_id = 0; const auto& [type, target, target_mode]: relationships_)
  {
    rel_id++;
    if(target_mode.empty())
    {
      xml_data += xml_empty_tag("Relationship", {
                                                  {"Id", std::format("rId{}", std::to_string(rel_id))},
                                                  {"Type", type},
                                                  {"Target", target},
      });
    }
    else
    {
      xml_data += xml_empty_tag("Relationship", {
                                                  {"Id", std::format("rId{}", std::to_string(rel_id))},
                                                  {"Type", type},
                                                  {"Target", target},
                                                  {"TargetMode", target_mode},
      });
    }
  }
  xml_data += xml_end_tag("Relationships");

  return xml_data;
}

void relationships_t::add(std::string_view schema, std::string_view type, std::string_view target,
                          std::string_view target_mode)
{
  if(schema.empty() || type.empty() || target.empty())
  {
    return;
  }

  relationships_.emplace_back(std::string(schema) + std::string(type), target, target_mode);
}

void relationships_t::add(std::string_view schema, std::string_view type, std::string_view target)
{
  if(schema.empty() || type.empty() || target.empty())
  {
    return;
  }

  relationships_.emplace_back(std::string(schema) + std::string(type), target, "");
}

}
