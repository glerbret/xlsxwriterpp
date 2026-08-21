/*
 * table - A libxlsxwriter library for creating Excel XLSX table files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#ifndef XWPP_TABLE_H
#define XWPP_TABLE_H

#include "xwpp/common.h"
#include "xwpp/worksheet.h"

#include <cstdint>

namespace xwpp
{

class table_t
{
public:
  explicit table_t(table_obj_t table_obj);

  [[nodiscard]] std::string assemble_xml_file() const;

private:
  [[nodiscard]] std::string write_table() const;
  [[nodiscard]] std::string write_auto_filter() const;
  [[nodiscard]] static std::string write_table_column(uint16_t id, const table_column_t& column);
  [[nodiscard]] std::string write_table_columns() const;
  [[nodiscard]] std::string write_table_style_info() const;

  table_obj_t table_obj_;
};

}

#endif
