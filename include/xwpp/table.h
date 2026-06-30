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
  table_t(const table_obj_t table_obj);

  [[nodiscard]] std::string assemble_xml_file();

private:
  [[nodiscard]] std::string write_table();
  [[nodiscard]] std::string write_auto_filter();
  [[nodiscard]] std::string write_table_column(uint16_t id, const table_column_t& column);
  [[nodiscard]] std::string write_table_columns();
  [[nodiscard]] std::string write_table_style_info();

  table_obj_t table_obj_;
};

}

#endif
