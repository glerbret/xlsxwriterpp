/*
 * shared_strings - A library for creating Excel XLSX sst files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/shared_strings.h"

#include "xwpp/xmlwriter.h"

#include <map>
#include <string>
#include <tuple>
#include <vector>

namespace xwpp
{

namespace
{

[[nodiscard]] std::string write_rich_si(const std::string& str)
{
  return xml_rich_si_element(str);
}

[[nodiscard]] std::string write_t(const std::string& str)
{
  std::vector<std::tuple<std::string, std::string>> attributes;
  // Add attribute to preserve leading or trailing whitespace.
  if(std::isspace(str.front()) != 0 || std::isspace(str.back()) != 0)
  {
    attributes.emplace_back("xml:space", "preserve");
  }

  return xml_data_element("t", str, attributes);
}

[[nodiscard]] std::string write_si(const std::string& str)
{
  const std::string encoded = escape_control_characters(str);

  std::string xml_data = xml_start_tag("si");
  xml_data += write_t(encoded);
  xml_data += xml_end_tag("si");

  return xml_data;
}

}

bool shared_strings_t::has_string() const
{
  return string_count_ != 0;
}

std::string shared_strings_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_sst();
  xml_data += write_sst_strings();
  xml_data += xml_end_tag("sst");

  return xml_data;
}

shared_strings_element_t shared_strings_t::get_index(const std::string& str, bool is_rich_string)
{
  // If element already exists, return it
  if(const auto it = rb_tree_.find(str); it != std::end(rb_tree_))
  {
    string_count_++;
    return order_list_[it->second];
  }

  const shared_strings_element_t elt{
    .index_          = unique_count_,
    .string_         = str,
    .is_rich_string_ = is_rich_string,
  };

  const size_t id = order_list_.size();
  order_list_.push_back(elt);
  rb_tree_[str] = id;
  string_count_++;
  unique_count_++;

  return elt;
}

std::string shared_strings_t::write_sst() const
{
  return xml_start_tag("sst", {
                                {"xmlns",       "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
                                {"count",       std::to_string(string_count_)                              },
                                {"uniqueCount", std::to_string(unique_count_)                              },
  });
}

std::string shared_strings_t::write_sst_strings() const
{
  std::string xml_data;

  for(const auto& elm: order_list_)
  {
    if(elm.is_rich_string_)
    {
      xml_data += write_rich_si(elm.string_);
    }
    else
    {
      xml_data += write_si(elm.string_);
    }
  }

  return xml_data;
}

}
