/*
 * relationships - A library for creating Excel XLSX relationships files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/relationships.h"

#include "xwpp/common.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

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
                                                    {"Id",     "rId" + std::to_string(rel_id)},
                                                    {"Type",   type                          },
                                                    {"Target", target                        },
      });
    }
    else
    {
      xml_data += xml_empty_tag("Relationship", {
                                                    {"Id",         "rId" + std::to_string(rel_id)},
                                                    {"Type",       type                          },
                                                    {"Target",     target                        },
                                                    {"TargetMode", target_mode                   },
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

void relationships_t::add_document(std::string_view type, std::string_view target)
{
  add(SCHEMA_DOCUMENT, type, target);
}

void relationships_t::add_package(std::string_view type, std::string_view target)
{
  add(SCHEMA_PACKAGE, type, target);
}

/// void
/// lxw_add_ms_package_relationship(lxw_relationships *self, const char *type,
///                                 const char *target)
/// {
///     _add_relationship(self, LXW_SCHEMA_MS, type, target, NULL);
/// }

void relationships_t::add_worksheet_relationship(std::string_view type, std::string_view target,
                                                 std::string_view target_mode)
{
  add(SCHEMA_DOCUMENT, type, target, target_mode);
}

/// void
/// lxw_add_rich_value_relationship(lxw_relationships *self)
/// {
///     _add_relationship(self,
///                       "http://schemas.microsoft.com/office/2022/10/relationships/",
///                       "richValueRel", "richData/richValueRel.xml", NULL);
///     _add_relationship(self,
///                       "http://schemas.microsoft.com/office/2017/06/relationships/",
///                       "rdRichValue", "richData/rdrichvalue.xml", NULL);
///     _add_relationship(self,
///                       "http://schemas.microsoft.com/office/2017/06/relationships/",
///                       "rdRichValueStructure",
///                       "richData/rdrichvaluestructure.xml", NULL);
///     _add_relationship(self,
///                       "http://schemas.microsoft.com/office/2017/06/relationships/",
///                       "rdRichValueTypes", "richData/rdRichValueTypes.xml",
///                       NULL);
/// }

}
