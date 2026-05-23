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
std::string datetime_to_iso8601_date(time_t* /*timer*/)
{
  // TODO Should be initialised in workbook constructor with local date and time
  const auto now = std::chrono::time_point_cast<std::chrono::seconds>(std::chrono::system_clock::now());
  /// struct tm *tmp_datetime;
  /// time_t current_time = time(NULL);

  /// if (*timer)
  ///   tmp_datetime = gmtime(timer);
  /// else
  ///   tmp_datetime = gmtime(&current_time);

  /// strftime(str, size - 1, "%Y-%m-%dT%H:%M:%SZ", tmp_datetime);
  return std::format("{:%FT%TZ}", now);
}

}

namespace xwpp
{

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
  /// if (self->properties->author) {
  ///   lxw_xml_data_element(self->file, "dc:creator",
  ///                             self->properties->author, NULL);
  ///   }
  ///     else {
  return xml_data_element("dc:creator");
  ///   }
}

std::string core_t::write_cp_last_modified_by() const
{
  /// if (self->properties->author) {
  ///   lxw_xml_data_element(self->file, "cp:lastModifiedBy", self->properties->author, NULL);
  /// }
  /// else {
  return xml_data_element("cp:lastModifiedBy");
  /// }
}

std::string core_t::write_dcterms_created() const
{
  return xml_data_element("dcterms:created", datetime_to_iso8601_date(nullptr /* TODO &self->properties->created*/),
                          {
                              {"xsi:type", "dcterms:W3CDTF"}
  });
}

std::string core_t::write_dcterms_modified() const
{
  return xml_data_element("dcterms:modified", datetime_to_iso8601_date(nullptr /* TODO &self->properties->created*/),
                          {
                              {"xsi:type", "dcterms:W3CDTF"}
  });
}

std::string core_t::write_dc_title() const
{
  /// if (!self->properties->title)
  ///   return;

  /// return xml_data_element("dc:title", self->properties->title,);
  return "";
}

std::string core_t::write_dc_subject() const
{
  /// if (!self->properties->subject)
  ///   return;

  /// return xml_data_element("dc:subject", self->properties->subject);
  return "";
}

std::string core_t::write_cp_keywords() const
{
  /// if (!self->properties->keywords)
  ///   return;

  /// return xml_data_element(self->file, "cp:keywords", self->properties->keywords);
  return "";
}

std::string core_t::write_dc_description() const
{
  /// if (!self->properties->comments)
  ///   return;

  ///  return xml_data_element("dc:description", self->properties->comments);
  return "";
}

std::string core_t::write_cp_category() const
{
  /// if (!self->properties->category)
  ///   return;

  /// return xml_data_element("cp:category", self->properties->category);
  return "";
}

std::string core_t::write_cp_content_status() const
{
  /// if (!self->properties->status)
  ///   return;

  /// return xml_data_element("cp:contentStatus", self->properties->status);
  return "";
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
