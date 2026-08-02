/*
 * rich_value_types - A library for creating Excel XLSX rich_value_types files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/rich_value_types.h"

#include "xwpp/utility.h"
#include "xwpp/xmlwriter.h"

namespace xwpp
{

std::string rich_value_types_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_rv_types_info();
  xml_data += write_key_flags();
  xml_data += xml_end_tag("rvTypesInfo");

  return xml_data;
}

std::string rich_value_types_t::write_rv_types_info() const
{
  return xml_start_tag("rvTypesInfo", {
                                        {"xmlns",        "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"},
                                        {"xmlns:mc",     "http://schemas.openxmlformats.org/markup-compatibility/2006"     },
                                        {"mc:Ignorable", "x"                                                               },
                                        {"xmlns:x",      "http://schemas.openxmlformats.org/spreadsheetml/2006/main"       },
  });
}

std::string rich_value_types_t::write_flag(const std::string& name) const
{
  return xml_empty_tag("flag", {
                                 {"name",  name},
                                 {"value", "1" },
  });
}

std::string rich_value_types_t::write_key(const std::string& name) const
{
  return xml_start_tag("key", {
                                {"name", name}
  });
}

std::string rich_value_types_t::write_key_flags() const
{
  std::vector<std::tuple<std::string, std::string, std::string>> key_flags{
    {"_Self",                   "ExcludeFromFile",           "ExcludeFromCalcComparison"},
    {"_DisplayString",          "ExcludeFromCalcComparison", ""                         },
    {"_Flags",                  "ExcludeFromCalcComparison", ""                         },
    {"_Format",                 "ExcludeFromCalcComparison", ""                         },
    {"_SubLabel",               "ExcludeFromCalcComparison", ""                         },
    {"_Attribution",            "ExcludeFromCalcComparison", ""                         },
    {"_Icon",                   "ExcludeFromCalcComparison", ""                         },
    {"_Display",                "ExcludeFromCalcComparison", ""                         },
    {"_CanonicalPropertyNames", "ExcludeFromCalcComparison", ""                         },
    {"_ClassificationId",       "ExcludeFromCalcComparison", ""                         },
  };

  std::string xml_data = xml_start_tag("global");
  xml_data += xml_start_tag("keyFlags");

  for(const auto& [key, flag1, flag2]: key_flags)
  {
    xml_data += write_key(key);
    xml_data += write_flag(flag1);
    if(!flag2.empty())
    {
      xml_data += write_flag(flag2);
    }
    xml_data += xml_end_tag("key");
  }
  xml_data += xml_end_tag("keyFlags");
  xml_data += xml_end_tag("global");

  return xml_data;
}

}
