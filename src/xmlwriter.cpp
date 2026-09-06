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

const std::string XML_AMP{"&amp;"};
const std::string XML_LT{"&lt;"};
const std::string XML_GT{"&gt;"};
const std::string XML_QUOT{"&quot;"};
const std::string XML_NL{"&#xA;"};

// Escape XML characters in data sections of tags.
// Note, this is different from _escape_attributes() in that double quotes are not escaped by Excel.
[[nodiscard]] std::string escape_data(std::string_view data)
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

// TODO A enlever
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

attributes_t::attributes_t(const std::vector<std::tuple<std::string, std::string>>& attributes)
  : attributes_{attributes}
{
}

void attributes_t::add_attribute(std::string_view key, std::string_view value)
{
  attributes_.emplace_back(key, value);
}

void attributes_t::add_attribute(std::string_view key, const char* value)
{
  attributes_.emplace_back(key, value);
}

void attributes_t::add_attribute(std::string_view key, bool value)
{
  attributes_.emplace_back(key, std::format("{:d}", value));
}

void attributes_t::add_attribute(std::string_view key, color_t value)
{
  attributes_.emplace_back(key, std::format("FF{:06X}", static_cast<uint32_t>(value) & COLOR_MASK));
}

std::string attributes_t::to_string() const
{
  std::string out;

  for(const auto& [key, value]: attributes_)
  {
    out += std::format(" {}=\"{}\"", key, escape_attribute_value(value));
  }

  return out;
}

bool attributes_t::empty() const
{
  return attributes_.empty();
}

[[nodiscard]] std::string attributes_t::escape_attribute_value(std::string_view attribute_value) const
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

std::string xml_declaration()
{
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
}

std::string xml_start_tag(std::string_view tag, const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{}{}>", tag, escaped_attributes(attributes));
}

std::string xml_start_tag(std::string_view tag, const attributes_t& attributes)
{
  return std::format("<{}{}>", tag, attributes.to_string());
}

std::string xml_start_tag(std::string_view tag)
{
  return std::format("<{}>", tag);
}

std::string xml_end_tag(std::string_view tag)
{
  return std::format("</{}>", tag);
}

std::string xml_empty_tag(std::string_view tag, const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{}{}/>", tag, escaped_attributes(attributes));
}

std::string xml_empty_tag(std::string_view tag, const attributes_t& attributes)
{
  return std::format("<{}{}/>", tag, attributes.to_string());
}

std::string xml_empty_tag(std::string_view tag)
{
  return std::format("<{}/>", tag);
}

std::string xml_data_element(std::string_view tag, std::string_view data,
                             const std::vector<std::tuple<std::string, std::string>>& attributes)
{
  return std::format("<{0}{1}>{2}</{0}>", tag, escaped_attributes(attributes), escape_data(data));
}

std::string xml_data_element(std::string_view tag, std::string_view data, const attributes_t& attributes)
{
  return std::format("<{0}{1}>{2}</{0}>", tag, attributes.to_string(), escape_data(data));
}

std::string xml_data_element(std::string_view tag, std::string_view data)
{
  return std::format("<{0}>{1}</{0}>", tag, escape_data(data));
}

std::string xml_data_element(std::string_view tag)
{
  return std::format("<{}/>", tag);
}

std::string xml_rich_si_element(std::string_view str)
{
  return std::format("<si>{}</si>", str);
}

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

  for(size_t i{0}; i < str.size(); i++)
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
        if(i >= str.size() - 2 || (isxdigit(str[i + 1]) == 0) || (isxdigit(str[i + 2]) == 0))
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

}
