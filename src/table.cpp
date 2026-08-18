/*
 * table - A library for creating Excel XLSX table files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xwpp/table.h"

#include "xwpp/utility.h"
#include "xwpp/worksheet.h"
#include "xwpp/xmlwriter.h"

#include <format>
#include <string>
#include <tuple>
#include <vector>

namespace xwpp
{

table_t::table_t(table_obj_t table_obj)
  : table_obj_{std::move(table_obj)}
{
}

std::string table_t::assemble_xml_file() const
{
  std::string xml_data = xml_declaration();
  xml_data += write_table();
  xml_data += write_auto_filter();
  xml_data += write_table_columns();
  xml_data += write_table_style_info();
  xml_data += xml_end_tag("table");

  return xml_data;
}

std::string table_t::write_table() const
{
  std::vector<std::tuple<std::string, std::string>> attributes{
    {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
    {"id",    std::to_string(table_obj_.id_)                             },
  };

  if(!table_obj_.name_.empty())
  {
    attributes.emplace_back("name", table_obj_.name_);
    attributes.emplace_back("displayName", table_obj_.name_);
  }
  else
  {
    attributes.emplace_back("name", "Table1");
    attributes.emplace_back("displayName", "Table1");
  }

  attributes.emplace_back("ref", table_obj_.sqref_);

  if(table_obj_.no_header_row_)
  {
    attributes.emplace_back("headerRowCount", "0");
  }

  if(table_obj_.total_row_)
  {
    attributes.emplace_back("totalsRowCount", "1");
  }
  else
  {
    attributes.emplace_back("totalsRowShown", "0");
  }

  return xml_start_tag("table", attributes);
}

std::string table_t::write_auto_filter() const
{
  if(table_obj_.no_autofilter_)
  {
    return "";
  }

  return xml_empty_tag("autoFilter", {
                                       {"ref", table_obj_.filter_sqref_}
  });
}

std::string table_t::write_table_column(uint16_t id, const table_column_t& column)
{
  std::vector<std::tuple<std::string, std::string>> attributes{
    {"id",   std::to_string(id)},
    {"name", column.header_    }
  };

  if(!column.total_string_.empty())
  {
    attributes.emplace_back("totalsRowLabel", column.total_string_);
  }
  else if(column.total_function_ != table_total_functions_t::NONE)
  {
    if(column.total_function_ == table_total_functions_t::AVERAGE)
    {
      attributes.emplace_back("totalsRowFunction", "average");
    }
    if(column.total_function_ == table_total_functions_t::COUNT_NUMS)
    {
      attributes.emplace_back("totalsRowFunction", "countNums");
    }
    if(column.total_function_ == table_total_functions_t::COUNT)
    {
      attributes.emplace_back("totalsRowFunction", "count");
    }
    if(column.total_function_ == table_total_functions_t::MAX)
    {
      attributes.emplace_back("totalsRowFunction", "max");
    }
    if(column.total_function_ == table_total_functions_t::MIN)
    {
      attributes.emplace_back("totalsRowFunction", "min");
    }
    if(column.total_function_ == table_total_functions_t::STD_DEV)
    {
      attributes.emplace_back("totalsRowFunction", "stdDev");
    }
    if(column.total_function_ == table_total_functions_t::SUM)
    {
      attributes.emplace_back("totalsRowFunction", "sum");
    }
    if(column.total_function_ == table_total_functions_t::VAR)
    {
      attributes.emplace_back("totalsRowFunction", "var");
    }
  }

  if(column.format_)
  {
    const int32_t dfx_id = column.format_->get_dxf_index_(column.format_);
    attributes.emplace_back("dataDxfId", std::to_string(dfx_id));
  }

  if(!column.formula_.empty())
  {
    std::string xml_data = xml_start_tag("tableColumn", attributes);
    xml_data += xml_data_element("calculatedColumnFormula", column.formula_);
    xml_data += xml_end_tag("tableColumn");

    return xml_data;
  }
  else
  {
    return xml_empty_tag("tableColumn", attributes);
  }
}

std::string table_t::write_table_columns() const
{
  std::string xml_data = xml_start_tag("tableColumns", {
                                                         {"count", std::to_string(table_obj_.columns_.size())}
  });
  for(uint16_t i = 0; const auto& column: table_obj_.columns_)
  {
    xml_data += write_table_column(i + 1, column);
    i++;
  }
  xml_data += xml_end_tag("tableColumns");

  return xml_data;
}

std::string table_t::write_table_style_info() const
{
  std::vector<std::tuple<std::string, std::string>> attributes;

  if(table_obj_.style_type_ == table_style_type_t::LIGHT)
  {
    if(table_obj_.style_type_number_ != 0)
    {
      attributes.emplace_back("name", std::format("TableStyleLight{}", table_obj_.style_type_number_));
    }
  }
  else if(table_obj_.style_type_ == table_style_type_t::MEDIUM)
  {
    attributes.emplace_back("name", std::format("TableStyleMedium{}", table_obj_.style_type_number_));
  }
  else if(table_obj_.style_type_ == table_style_type_t::DARK)
  {
    attributes.emplace_back("name", std::format("TableStyleDark{}", table_obj_.style_type_number_));
  }
  else
  {
    attributes.emplace_back("name", "TableStyleMedium9");
  }

  if(table_obj_.first_column_)
  {
    attributes.emplace_back("showFirstColumn", "1");
  }
  else
  {
    attributes.emplace_back("showFirstColumn", "0");
  }

  if(table_obj_.last_column_)
  {
    attributes.emplace_back("showLastColumn", "1");
  }
  else
  {
    attributes.emplace_back("showLastColumn", "0");
  }

  if(table_obj_.no_banded_rows_)
  {
    attributes.emplace_back("showRowStripes", "0");
  }
  else
  {
    attributes.emplace_back("showRowStripes", "1");
  }

  if(table_obj_.banded_columns_)
  {
    attributes.emplace_back("showColumnStripes", "1");
  }
  else
  {
    attributes.emplace_back("showColumnStripes", "0");
  }

  return xml_empty_tag("tableStyleInfo", attributes);
}

}
