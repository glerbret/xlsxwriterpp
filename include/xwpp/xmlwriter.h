/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
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
 */

#ifndef XWPP_XMLWRITER_H
#define XWPP_XMLWRITER_H

#include <string>
#include <string_view>
#include <tuple>
#include <vector>

// TODO Create type for attributes with support of different types (string, integer float, ...)

namespace xwpp
{

const size_t MAX_ATTRIBUTE_LENGTH = 2080;

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
[[nodiscard]] std::string xml_start_tag(std::string_view tag);

/**
 * Write an XML start tag with optional un-encoded attributes.
 * This is a minor optimization for attributes that don't need encoding.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
/// void lxw_xml_start_tag_unencoded(FILE *xmlfile,
///                                  const char *tag,
///                                  struct xml_attribute_list *attributes);

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
[[nodiscard]] std::string xml_empty_tag(std::string_view tag);

/**
 * Write an XML empty tag with optional un-encoded attributes.
 * This is a minor optimization for attributes that don't need encoding.
 *
 * @param xmlfile    A FILE pointer to the output XML file.
 * @param tag        The XML tag to write.
 * @param attributes An optional list of attributes to add to the tag.
 */
/// void lxw_xml_empty_tag_unencoded(FILE *xmlfile,
///                                  const char *tag,
///                                  struct xml_attribute_list *attributes);

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
[[nodiscard]] std::string xml_data_element(std::string_view tag, std::string_view data);
[[nodiscard]] std::string xml_data_element(std::string_view tag);

[[nodiscard]] std::string xml_rich_si_element(std::string_view str);

/// uint8_t lxw_has_control_characters(const char *string);
[[nodiscard]] std::string escape_control_characters(std::string_view str);
[[nodiscard]] std::string escape_url_characters(const std::string& str, bool escape_hash);

[[nodiscard]] std::string escape_data(std::string_view data);

}

#endif
