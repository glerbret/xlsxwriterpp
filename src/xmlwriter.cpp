/*
 * xmlwriter - A base library for Xlsxwriter++ libraries.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/xmlwriter.h"

#include <cctype>
#include <format>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

namespace xwpp
{

namespace
{

const std::string XML_AMP  = "&amp;";
const std::string XML_LT   = "&lt;";
const std::string XML_GT   = "&gt;";
const std::string XML_QUOT = "&quot;";
const std::string XML_NL   = "&#xA;";

std::string escape_attribute_value(std::string_view attribute_value)
{
  std::string encoded;
  encoded.reserve(attribute_value.size() * 2);

  for(auto c: attribute_value)
  {
    switch(c)
    {
      case '&':
        encoded += XML_AMP;
        break;

      case '<':
        encoded += XML_LT;
        break;

      case '>':
        encoded += XML_GT;
        break;

      case '"':
        encoded += XML_QUOT;
        break;

      case '\n':
        encoded += XML_NL;
        break;

      default:
        encoded.push_back(c);
    }
  }

  return encoded;
}

std::string escaped_attributes(const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  std::string out;

  for(const auto& [key, value]: attributes)
  {
    out += std::format(" {}=\"{}\"", key, escape_attribute_value(value));
  }

  return out;
}

}

std::string xml_declaration()
{
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
}

std::string xml_start_tag(std::string_view tag, const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{}{}>", tag, escaped_attributes(attributes));
}

std::string xml_start_tag(std::string_view tag)
{
  return std::format("<{}>", tag);
}

/// void lxw_xml_start_tag_unencoded(FILE *xmlfile,
///                             const char *tag,
///                             struct xml_attribute_list *attributes)
/// {
///     struct xml_attribute *attribute;

///     fprintf(xmlfile, "<%s", tag);

///     if (attributes) {
///         STAILQ_FOREACH(attribute, attributes, list_entries) {
///             fprintf(xmlfile, " %s=\"%s\"", attribute->key, attribute->value);
///         }
///     }

///     fprintf(xmlfile, ">");
/// }

std::string xml_end_tag(std::string_view tag)
{
  return std::format("</{}>", tag);
}

std::string xml_empty_tag(std::string_view tag, const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{}{}/>", tag, escaped_attributes(attributes));
}

std::string xml_empty_tag(std::string_view tag)
{
  return std::format("<{}/>", tag);
}

/// void lxw_xml_empty_tag_unencoded(FILE *xmlfile,
///                             const char *tag,
///                             struct xml_attribute_list *attributes)
/// {
///     struct xml_attribute *attribute;

///     fprintf(xmlfile, "<%s", tag);

///     if (attributes) {
///         STAILQ_FOREACH(attribute, attributes, list_entries) {
///             fprintf(xmlfile, " %s=\"%s\"", attribute->key, attribute->value);
///         }
///     }

///     fprintf(xmlfile, "/>");
/// }

std::string xml_data_element(std::string_view tag, std::string_view data,
                             const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{0}{1}>{2}</{0}>", tag, escaped_attributes(attributes), escape_data(data));
}
std::string xml_data_element(std::string_view tag, std::string_view data)
{
  return std::format("<{0}>{1}</{0}>", tag, escape_data(data));
}
std::string xml_data_element(std::string_view tag)
{
  return std::format("<{0}/>", tag);
}

std::string xml_rich_si_element(std::string_view str)
{
  return std::format("<si>{}</si>", str);
}

// Escape XML characters in data sections of tags.
// Note, this is different from _escape_attributes() in that double quotes are not escaped by Excel.
std::string escape_data(std::string_view data)
{
  std::string encoded;
  encoded.reserve(data.size() * 2);

  for(auto c: data)
  {
    switch(c)
    {
      case '&':
        encoded += XML_AMP;
        break;

      case '<':
        encoded += XML_LT;
        break;

      case '>':
        encoded += XML_GT;
        break;

      default:
        encoded.push_back(c);
    }
  }

  return encoded;
}

/// uint8_t lxw_has_control_characters(const char *string)
/// {
///     while (*string) {
///         /* 0xE0 == 0b11100000 masks values > 0x19 == 0b00011111. */
///         if (!(*string & 0xE0) && *string != 0x0A && *string != 0x09)
///             return LXW_TRUE;

///         string++;
///     }

///     return LXW_FALSE;
/// }

// Escape control characters in strings with _xHHHH_.
std::string escape_control_characters(std::string_view str)
{
  std::string encoded;
  encoded.reserve(str.size() * 2);

  for(auto c: str)
  {
    switch(c)
    {
      case '\x01':
      case '\x02':
      case '\x03':
      case '\x04':
      case '\x05':
      case '\x06':
      case '\x07':
      case '\x08':
      case '\x0B':
      case '\x0C':
      case '\x0D':
      case '\x0E':
      case '\x0F':
      case '\x10':
      case '\x11':
      case '\x12':
      case '\x13':
      case '\x14':
      case '\x15':
      case '\x16':
      case '\x17':
      case '\x18':
      case '\x19':
      case '\x1A':
      case '\x1B':
      case '\x1C':
      case '\x1D':
      case '\x1E':
      case '\x1F':
        encoded += std::format("_x{:04X}_", c);
        break;
      default:
        encoded.push_back(c);
    }
  }

  return encoded;
}

std::string escape_url_characters(const std::string& str, bool escape_hash)
{
  std::string encoded;
  encoded.reserve(2 * str.size());

  for(size_t i = 0; i < str.size(); i++)
  {
    switch(str[i])
    {
      case ' ':
      case '"':
      case '<':
      case '>':
      case '[':
      case ']':
      case '`':
      case '^':
      case '{':
      case '}':
        encoded += std::format("%{:02x}", str[i]);
        break;

      case '#':
        // This is only escaped for "external:" style links.
        if(escape_hash)
        {
          encoded += std::format("%{:02x}", str[i]);
        }
        else
        {
          encoded.push_back(str[i]);
        }
        break;

      case '%':
        // Only escape % if it isn't already an escape.
        if(i >= str.size() - 2 || !isxdigit(str[i + 1]) || !isxdigit(str[i + 1]))
        {
          encoded += std::format("%{:02x}", str[i]);
        }
        else
        {
          encoded.push_back(str[i]);
        }
        break;

      default:
        encoded.push_back(str[i]);
        break;
    }
  }

  return encoded;
}

/// struct xml_attribute * lxw_new_attribute_str(const char *key, const char *value)
/// {
///     struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

///     LXW_ATTRIBUTE_COPY(attribute->key, key);
///     LXW_ATTRIBUTE_COPY(attribute->value, value);

///     return attribute;
/// }

/// struct xml_attribute * lxw_new_attribute_int(const char *key, int32_t value)
/// {
///     struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

///     LXW_ATTRIBUTE_COPY(attribute->key, key);
///     lxw_snprintf(attribute->value, LXW_MAX_ATTRIBUTE_LENGTH, "%d", value);

///     return attribute;
/// }

/// struct xml_attribute * lxw_new_attribute_dbl(const char *key, double value)
/// {
///     struct xml_attribute *attribute = malloc(sizeof(struct xml_attribute));

///     LXW_ATTRIBUTE_COPY(attribute->key, key);
///     lxw_sprintf_dbl(attribute->value, value);

///     return attribute;
/// }

}
