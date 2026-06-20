/*
 * app - A library for creating Excel XLSX app files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xwpp/app.h"

#include "xwpp/common.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

std::string app_t::write_properties() const
{
  return xml_start_tag("Properties", {
                                         {"xmlns",    SCHEMA_OFFICEDOC + "/extended-properties"},
                                         {"xmlns:vt", SCHEMA_OFFICEDOC + "/docPropsVTypes"     },
  });
}

std::string app_t::write_application() const
{
  // TODO Check why name is important and which default are different following the case (e.g. chart_data_labels)
  // Application name must be "Microsoft Excel", even some default change
  return xml_data_element("Application", "Microsoft Excel");
}

std::string app_t::write_doc_security() const
{
  if(doc_security_ == 2)
  {
    return xml_data_element("DocSecurity", "2");
  }
  else
  {
    return xml_data_element("DocSecurity", "0");
  }
}

std::string app_t::write_scale_crop() const
{
  return xml_data_element("ScaleCrop", "false");
}

std::string app_t::write_vt_lpstr(const std::string& str) const
{
  return xml_data_element("vt:lpstr", str);
}

std::string app_t::write_vt_i4(const std::string& value) const
{
  return xml_data_element("vt:i4", value);
}

std::string app_t::write_vt_variant(const std::string& key, const std::string& value) const
{
  // Write the vt:lpstr element.
  std::string xml_data = xml_start_tag("vt:variant");
  xml_data += write_vt_lpstr(key);
  xml_data += xml_end_tag("vt:variant");

  // Write the vt:i4 element.
  xml_data += xml_start_tag("vt:variant");
  xml_data += write_vt_i4(value);
  xml_data += xml_end_tag("vt:variant");

  return xml_data;
}

std::string app_t::write_vt_vector_heading_pairs() const
{
  std::string xml_data = xml_start_tag("vt:vector", {
                                                        {"size",     std::to_string(heading_pairs_.size() * 2)},
                                                        {"baseType", "variant"                                },
  });

  for(const auto& [key, value]: heading_pairs_)
  {
    xml_data += write_vt_variant(key, value);
  }

  xml_data += xml_end_tag("vt:vector");
  return xml_data;
}

std::string app_t::write_vt_vector_lpstr_named_parts() const
{
  std::string xml_data = xml_start_tag("vt:vector", {
                                                        {"size",     std::to_string(part_names_.size())},
                                                        {"baseType", "lpstr"                           },
  });

  for(const auto& part_name: part_names_)
  {
    xml_data += write_vt_lpstr(part_name.name_);
  }
  xml_data += xml_end_tag("vt:vector");

  return xml_data;
}

std::string app_t::write_heading_pairs() const
{
  std::string xml_data = xml_start_tag("HeadingPairs");

  // Write the vt:vector element.
  xml_data += write_vt_vector_heading_pairs();

  xml_data += xml_end_tag("HeadingPairs");
  return xml_data;
}

std::string app_t::write_titles_of_parts() const
{
  std::string xml_data = xml_start_tag("TitlesOfParts");

  // Write the vt:vector element.
  xml_data += write_vt_vector_lpstr_named_parts();

  xml_data += xml_end_tag("TitlesOfParts");
  return xml_data;
}

std::string app_t::write_manager() const
{
  if(properties_.manager_.empty())
  {
    return "";
  }
  return xml_data_element("Manager", properties_.manager_);
}

std::string app_t::write_company() const
{
  if(!properties_.company_.empty())
  {
    return xml_data_element("Company", properties_.company_);
  }
  else
  {
    return xml_data_element("Company", "");
  }
}

std::string app_t::write_links_up_to_date() const
{
  return xml_data_element("LinksUpToDate", "false");
}

std::string app_t::write_shared_doc() const
{
  return xml_data_element("SharedDoc", "false");
}

std::string app_t::write_hyperlink_base() const
{
  /// lxw_doc_properties *properties = self->properties;

  /// if (!properties)
  ///   return;

  /// if (properties->hyperlink_base)
  ///   lxw_xml_data_element(self->file, "HyperlinkBase", properties->hyperlink_base, NULL);
  return "";
}

std::string app_t::write_hyperlinks_changed() const
{
  return xml_data_element("HyperlinksChanged", "false");
}

std::string app_t::write_app_version() const
{
  return xml_data_element("AppVersion", "12.0000");
}

std::string app_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_properties();
  xml_data += write_application();
  xml_data += write_doc_security();
  xml_data += write_scale_crop();
  xml_data += write_heading_pairs();
  xml_data += write_titles_of_parts();
  xml_data += write_manager();
  xml_data += write_company();
  xml_data += write_links_up_to_date();
  xml_data += write_shared_doc();
  xml_data += write_hyperlink_base();
  xml_data += write_hyperlinks_changed();
  xml_data += write_app_version();
  xml_data += xml_end_tag("Properties");
  return xml_data;
}

void app_t::add_part_name(const std::string& name)
{
  if(name.empty())
  {
    return;
  }

  part_names_.emplace_back(name);
}

void app_t::add_heading_pair(const std::string& key, const std::string& value)
{
  if(key.empty() || value.empty())
  {
    return;
  }

  heading_pairs_.emplace_back(key, value);
}

void app_t::set_properties(const doc_properties_t& properties)
{
  properties_ = properties;
}

}
