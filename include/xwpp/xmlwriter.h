/*
 * xmlwriter - A Xlsxwriter++ library for creating Excel XLSX
 *             XML files.
 *
 * The xmlwriter library is used to create the XML sub-components files
 * in the Excel XLSX file format.
 *
 * This library is used in preference to a more generic XML library to allow
 * for customization and optimization for the XLSX file format.
 *
 * The xmlwriter functions are only used internally and do not need to be
 * called directly by the end user.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_XMLWRITER_H
#define XWPP_XMLWRITER_H

#include <concepts>
#include <format>
#include <string>
#include <string_view>
#include <tuple>
#include <type_traits>
#include <vector>

#include "format.h"

namespace xwpp
{

class attributes_t
{
public:
  attributes_t() = default;
  explicit attributes_t(const std::vector<std::tuple<std::string, std::string>>& attributes);

  void add_attribute(std::string_view key, std::string_view value);
  void add_attribute(std::string_view key, const char* value);
  void add_attribute(std::string_view key, bool value);
  template<std::integral T>
  void add_attribute(std::string_view key, T value)
  {
    attributes_.emplace_back(key, std::to_string(value));
  }
  template<std::floating_point T>
  void add_attribute(std::string_view key, T value)
  {
    attributes_.emplace_back(key, std::format("{}", value));
  }
  void add_attribute(std::string_view key, color_t value);

  [[nodiscard]] std::string to_string() const;
  [[nodiscard]] bool empty() const;

private:
  std::vector<std::tuple<std::string, std::string>> attributes_;

  [[nodiscard]] std::string escape_attribute_value(std::string_view attribute_value) const;
};


/**
 * Create the XML declaration in an XML file.
 *
 * @return XML declaration.
 */
[[nodiscard]] std::string xml_declaration();

/**
 * Write an XML start tag with optional attributes.
 *
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 *
 * @return XML content.
 */
[[nodiscard]] std::string xml_start_tag(std::string_view tag,
                                        const std::vector<std::tuple<std::string, std::string>>& attributes);
[[nodiscard]] std::string xml_start_tag(std::string_view tag,
                                        const attributes_t& attributes);
[[nodiscard]] std::string xml_start_tag(std::string_view tag);

/**
 * Write an XML end tag.
 *
 * @param tag        The XML tag to write.
 *
 * @return XML content.
 */
[[nodiscard]] std::string xml_end_tag(std::string_view tag);

/**
 * Write an XML empty tag with optional attributes.
 *
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 *
 * @return XML content.
 */
[[nodiscard]] std::string xml_empty_tag(std::string_view tag,
                                        const std::vector<std::tuple<std::string, std::string>>& attributes);
[[nodiscard]] std::string xml_empty_tag(std::string_view tag, const attributes_t& attributes);
[[nodiscard]] std::string xml_empty_tag(std::string_view tag);

/**
 * Write an XML element containing data and optional attributes.
 *
 * @param tag        The XML tag to write.
 * @param data       The data section of the XML element.
 * @param attributes An optional list of attributes to add to the tag.
 *
 * @return XML content.
 */
[[nodiscard]] std::string xml_data_element(std::string_view tag, std::string_view data,
                                           const std::vector<std::tuple<std::string, std::string>>& attributes);
[[nodiscard]] std::string xml_data_element(std::string_view tag, std::string_view data,
                                           const attributes_t& attributes);
[[nodiscard]] std::string xml_data_element(std::string_view tag, std::string_view data);
[[nodiscard]] std::string xml_data_element(std::string_view tag);
[[nodiscard]] std::string xml_rich_si_element(std::string_view str);
[[nodiscard]] std::string escape_control_characters(std::string_view str);
[[nodiscard]] std::string escape_url_characters(const std::string& str, bool escape_hash);

}

#endif
