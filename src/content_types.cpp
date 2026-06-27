/*
 * content_types - A library for creating Excel XLSX content_types files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/content_types.h"

#include "xwpp/common.h"
#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

#include <string>

namespace xwpp
{

content_types_t::content_types_t()
{
  add_default("rels", APP_PACKAGE + "relationships+xml");
  add_default("xml", "application/xml");

  add_override("/docProps/app.xml", APP_DOCUMENT + "extended-properties+xml");
  add_override("/docProps/core.xml", APP_PACKAGE + "core-properties+xml");
  add_override("/xl/styles.xml", APP_DOCUMENT + "spreadsheetml.styles+xml");
  add_override("/xl/theme/theme1.xml", APP_DOCUMENT + "theme+xml");
}

void content_types_t::add_default(std::string_view key, std::string_view value)
{
  if(key.empty() || value.empty())
  {
    return;
  }

  default_types_.emplace_back(key, value);
}

void content_types_t::add_override(std::string_view key, std::string_view value)
{
  if(key.empty() || value.empty())
  {
    return;
  }

  overrides_.emplace_back(key, value);
}

void content_types_t::add_worksheet_name(std::string_view name)
{
  add_override(name, APP_DOCUMENT + "spreadsheetml.worksheet+xml");
}

std::string content_types_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += xml_start_tag("Types", {
                                         {"xmlns", SCHEMA_CONTENT}
  });

  for(const auto& [key, value]: default_types_)
  {
    xml_data += xml_empty_tag("Default", {
                                             {"Extension",   key  },
                                             {"ContentType", value}
    });
  }

  for(const auto& [key, value]: overrides_)
  {
    xml_data += xml_empty_tag("Override", {
                                              {"PartName",    key  },
                                              {"ContentType", value}
    });
  }
  xml_data += xml_end_tag("Types");
  return xml_data;
}

/// void _write_types(lxw_content_types *self)
/// {
/// struct xml_attribute_list attributes;
/// struct xml_attribute *attribute;

/// LXW_INIT_ATTRIBUTES();
/// LXW_PUSH_ATTRIBUTES_STR("xmlns", LXW_SCHEMA_CONTENT);

/// lxw_xml_start_tag(self->file, "Types", &attributes);

/// LXW_FREE_ATTRIBUTES();
/// }

void content_types_t::add_chartsheet_name(std::string_view name)
{
  add_override(name, APP_DOCUMENT + "spreadsheetml.chartsheet+xml");
}

void content_types_t::add_chart_name(std::string_view name)
{
  add_override(name, APP_DOCUMENT + "drawingml.chart+xml");
}

void content_types_t::add_drawing_name(std::string_view name)
{
  add_override(name, APP_DOCUMENT + "drawing+xml");
}

/// void lxw_ct_add_table_name(lxw_content_types *self, const char *name)
/// {
///   lxw_ct_add_override(self, name, LXW_APP_DOCUMENT "spreadsheetml.table+xml");
/// }

void content_types_t::add_vml_name()
{
  add_default("vml", APP_DOCUMENT + "vmlDrawing");
}

void content_types_t::add_comment_name(std::string_view name)
{
  add_override(name, APP_DOCUMENT + "spreadsheetml.comments+xml");
}

void content_types_t::add_shared_strings()
{
  add_override("/xl/sharedStrings.xml", APP_DOCUMENT + "spreadsheetml.sharedStrings+xml");
}

/// void lxw_ct_add_calc_chain(lxw_content_types *self)
/// {
///   lxw_ct_add_override(self, "/xl/calcChain.xml", LXW_APP_DOCUMENT "spreadsheetml.calcChain+xml");
/// }

void content_types_t::add_custom_properties()
{
  add_override("/docProps/custom.xml", APP_DOCUMENT + "custom-properties+xml");
}

void content_types_t::add_metadata()
{
  add_override("/xl/metadata.xml", APP_DOCUMENT + "spreadsheetml.sheetMetadata+xml");
}

void content_types_t::add_rich_value()
{
  add_override("/xl/richData/rdRichValueTypes.xml", APP_MSEXCEL + "rdrichvaluetypes+xml");
  add_override("/xl/richData/rdrichvalue.xml", APP_MSEXCEL + "rdrichvalue+xml");
  add_override("/xl/richData/rdrichvaluestructure.xml", APP_MSEXCEL + "rdrichvaluestructure+xml");
  add_override("/xl/richData/richValueRel.xml", APP_MSEXCEL + "richvaluerel+xml");
}

const std::string content_types_t::APP_PACKAGE{"application/vnd.openxmlformats-package."};
const std::string content_types_t::APP_DOCUMENT{"application/vnd.openxmlformats-officedocument."};
const std::string content_types_t::APP_MSEXCEL{"application/vnd.ms-excel."};

}
