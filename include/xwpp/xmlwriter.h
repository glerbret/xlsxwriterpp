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

/// #define LXW_ATTR_32              32

/// #define LXW_ATTRIBUTE_COPY(dst, src)                    \
///     do{                                                 \
///         strncpy(dst, src, LXW_MAX_ATTRIBUTE_LENGTH -1); \
///         dst[LXW_MAX_ATTRIBUTE_LENGTH - 1] = '\0';       \
///     } while (0)

/* Attribute used in XML elements. */
/// struct xml_attribute {
///     char key[LXW_MAX_ATTRIBUTE_LENGTH];
///     char value[LXW_MAX_ATTRIBUTE_LENGTH];

/* Make the struct a queue.h list element. */
///     STAILQ_ENTRY (xml_attribute) list_entries;
/// };

/* Use queue.h macros to define the xml_attribute_list type. */
/// STAILQ_HEAD(xml_attribute_list, xml_attribute);

/* Create a new attribute struct to add to a xml_attribute_list. */
/// struct xml_attribute *lxw_new_attribute_str(const char *key,
///                                             const char *value);
/// struct xml_attribute *lxw_new_attribute_int(const char *key, int32_t value);
/// struct xml_attribute *lxw_new_attribute_dbl(const char *key, double value);

/* Macro to initialize the xml_attribute_list pointers. */
/// #define LXW_INIT_ATTRIBUTES()                                 \
///     STAILQ_INIT(&attributes)

/* Macro to add attribute string elements to xml_attribute_list. */
/// #define LXW_PUSH_ATTRIBUTES_STR(key, value)                   \
///     do {                                                      \
///     attribute = lxw_new_attribute_str((key), (value));        \
///     STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
///     } while (0)

/* Macro to add attribute int values to xml_attribute_list. */
/// #define LXW_PUSH_ATTRIBUTES_INT(key, value)                   \
///     do {                                                      \
///     attribute = lxw_new_attribute_int((key), (value));        \
///     STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
///     } while (0)

/* Macro to add attribute double values to xml_attribute_list. */
/// #define LXW_PUSH_ATTRIBUTES_DBL(key, value)                   \
///     do {                                                      \
///     attribute = lxw_new_attribute_dbl((key), (value));        \
///     STAILQ_INSERT_TAIL(&attributes, attribute, list_entries); \
///     } while (0)

/* Macro to free xml_attribute_list and attribute. */
/// #define LXW_FREE_ATTRIBUTES()                                 \
///     do {                                                      \
///         while (!STAILQ_EMPTY(&attributes)) {                  \
///             attribute = STAILQ_FIRST(&attributes);            \
///             STAILQ_REMOVE_HEAD(&attributes, list_entries);    \
///             free(attribute);                                  \
///         }                                                     \
///     } while (0)

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
