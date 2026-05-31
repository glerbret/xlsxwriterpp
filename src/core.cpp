/*****************************************************************************
 * core - A library for creating Excel XLSX core files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/core.h"

#include "xwpp/xmlwriter.h"

#include <chrono>
#include <format>

namespace
{

// TODO Should be moved to utility
std::string datetime_to_iso8601_date(const std::chrono::system_clock::time_point& time)
{
  if(time.time_since_epoch().count() == 0)
  {
    return std::format("{:%FT%TZ}",
                       std::chrono::time_point_cast<std::chrono::seconds>(std::chrono::system_clock::now()));
  }

  return std::format("{:%FT%TZ}", std::chrono::time_point_cast<std::chrono::seconds>(time));
}

}

namespace xwpp
{

core_t::core_t(doc_properties_t properties)
  : properties_{std::move(properties)}
{
}

std::string core_t::write_cp_core_properties() const
{
  return xml_start_tag("cp:coreProperties",
                       {
                           {"xmlns:cp",       "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
                           {"xmlns:dc",       "http://purl.org/dc/elements/1.1/"                                       },
                           {"xmlns:dcterms",  "http://purl.org/dc/terms/"                                              },
                           {"xmlns:dcmitype", "http://purl.org/dc/dcmitype/"                                           },
                           {"xmlns:xsi",      "http://www.w3.org/2001/XMLSchema-instance"                              },
  });
}

std::string core_t::write_dc_creator() const
{
  if(!properties_.author_.empty())
  {
    return xml_data_element("dc:creator", properties_.author_);
  }
  else
  {
    return xml_data_element("dc:creator");
  }
}

std::string core_t::write_cp_last_modified_by() const
{
  if(!properties_.modif_author_.empty())
  {
    return xml_data_element("cp:lastModifiedBy", properties_.modif_author_);
  }
  else
  {
    return xml_data_element("cp:lastModifiedBy");
  }
}

std::string core_t::write_dcterms_created() const
{
  return xml_data_element("dcterms:created", datetime_to_iso8601_date(properties_.created_),
                          {
                              {"xsi:type", "dcterms:W3CDTF"}
  });
}

std::string core_t::write_dcterms_modified() const
{
  if(properties_.modified_.time_since_epoch().count() == 0)
  {
    return "";
  }

  return xml_data_element("dcterms:modified", datetime_to_iso8601_date(properties_.modified_),
                          {
                              {"xsi:type", "dcterms:W3CDTF"}
  });
}

std::string core_t::write_dc_title() const
{
  if(properties_.title_.empty())
  {
    return "";
  }

  return xml_data_element("dc:title", properties_.title_);
}

std::string core_t::write_dc_subject() const
{
  if(properties_.subject_.empty())
  {
    return "";
  }

  return xml_data_element("dc:subject", properties_.subject_);
}

std::string core_t::write_cp_keywords() const
{
  if(properties_.keywords_.empty())
  {
    return "";
  }

  return xml_data_element("cp:keywords", properties_.keywords_);
}

std::string core_t::write_dc_description() const
{
  if(properties_.comments_.empty())
  {
    return "";
  }

  return xml_data_element("dc:description", properties_.comments_);
}

std::string core_t::write_cp_category() const
{
  if(properties_.category_.empty())
  {
    return "";
  }

  return xml_data_element("cp:category", properties_.category_);
}

std::string core_t::write_cp_content_status() const
{
  if(properties_.status_.empty())
  {
    return "";
  }

  return xml_data_element("cp:contentStatus", properties_.status_);
}

std::string core_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();

  xml_data += write_cp_core_properties();
  xml_data += write_dc_title();
  xml_data += write_dc_subject();
  xml_data += write_dc_creator();
  xml_data += write_cp_keywords();
  xml_data += write_dc_description();
  xml_data += write_cp_last_modified_by();
  xml_data += write_dcterms_created();
  xml_data += write_dcterms_modified();
  xml_data += write_cp_category();
  xml_data += write_cp_content_status();
  xml_data += xml_end_tag("cp:coreProperties");

  return xml_data;
}

}
