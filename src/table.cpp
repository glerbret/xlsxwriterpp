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
  attributes_t attributes{{
    {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
    {"id",    std::to_string(table_obj_.id_)                             },
  }};

  if(!table_obj_.name_.empty())
  {
    attributes.add_attribute("name", table_obj_.name_);
    attributes.add_attribute("displayName", table_obj_.name_);
  }
  else
  {
    attributes.add_attribute("name", "Table1");
    attributes.add_attribute("displayName", "Table1");
  }

  attributes.add_attribute("ref", table_obj_.sqref_);

  if(table_obj_.no_header_row_)
  {
    attributes.add_attribute("headerRowCount", "0");
  }

  if(table_obj_.total_row_)
  {
    attributes.add_attribute("totalsRowCount", "1");
  }
  else
  {
    attributes.add_attribute("totalsRowShown", "0");
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
  attributes_t attributes{{
    {"id",   std::to_string(id)},
    {"name", column.header_    }
  }};

  if(!column.total_string_.empty())
  {
    attributes.add_attribute("totalsRowLabel", column.total_string_);
  }
  else if(column.total_function_ != table_total_functions_t::NONE)
  {
    if(column.total_function_ == table_total_functions_t::AVERAGE)
    {
      attributes.add_attribute("totalsRowFunction", "average");
    }
    if(column.total_function_ == table_total_functions_t::COUNT_NUMS)
    {
      attributes.add_attribute("totalsRowFunction", "countNums");
    }
    if(column.total_function_ == table_total_functions_t::COUNT)
    {
      attributes.add_attribute("totalsRowFunction", "count");
    }
    if(column.total_function_ == table_total_functions_t::MAX)
    {
      attributes.add_attribute("totalsRowFunction", "max");
    }
    if(column.total_function_ == table_total_functions_t::MIN)
    {
      attributes.add_attribute("totalsRowFunction", "min");
    }
    if(column.total_function_ == table_total_functions_t::STD_DEV)
    {
      attributes.add_attribute("totalsRowFunction", "stdDev");
    }
    if(column.total_function_ == table_total_functions_t::SUM)
    {
      attributes.add_attribute("totalsRowFunction", "sum");
    }
    if(column.total_function_ == table_total_functions_t::VAR)
    {
      attributes.add_attribute("totalsRowFunction", "var");
    }
  }

  if(column.format_)
  {
    const int32_t dfx_id = column.format_->get_dxf_index_(column.format_);
    attributes.add_attribute("dataDxfId", dfx_id);
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
  for(uint16_t i{0}; const auto& column: table_obj_.columns_)
  {
    xml_data += write_table_column(i + 1, column);
    i++;
  }
  xml_data += xml_end_tag("tableColumns");

  return xml_data;
}

std::string table_t::write_table_style_info() const
{
  attributes_t attributes;

  if(table_obj_.style_type_ == table_style_type_t::LIGHT)
  {
    if(table_obj_.style_type_number_ != 0)
    {
      attributes.add_attribute("name", std::format("TableStyleLight{}", table_obj_.style_type_number_));
    }
  }
  else if(table_obj_.style_type_ == table_style_type_t::MEDIUM)
  {
    attributes.add_attribute("name", std::format("TableStyleMedium{}", table_obj_.style_type_number_));
  }
  else if(table_obj_.style_type_ == table_style_type_t::DARK)
  {
    attributes.add_attribute("name", std::format("TableStyleDark{}", table_obj_.style_type_number_));
  }
  else
  {
    attributes.add_attribute("name", "TableStyleMedium9");
  }

  if(table_obj_.first_column_)
  {
    attributes.add_attribute("showFirstColumn", "1");
  }
  else
  {
    attributes.add_attribute("showFirstColumn", "0");
  }

  if(table_obj_.last_column_)
  {
    attributes.add_attribute("showLastColumn", "1");
  }
  else
  {
    attributes.add_attribute("showLastColumn", "0");
  }

  if(table_obj_.no_banded_rows_)
  {
    attributes.add_attribute("showRowStripes", "0");
  }
  else
  {
    attributes.add_attribute("showRowStripes", "1");
  }

  if(table_obj_.banded_columns_)
  {
    attributes.add_attribute("showColumnStripes", "1");
  }
  else
  {
    attributes.add_attribute("showColumnStripes", "0");
  }

  return xml_empty_tag("tableStyleInfo", attributes);
}

}
