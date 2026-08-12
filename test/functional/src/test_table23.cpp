/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_column(COLS("B:F"), 10.288);

  worksheet.write_string(CELL("A1"), "Column1");
  worksheet.write_string(CELL("F1"), "Total");
  worksheet.write_string(CELL("B1"), "Column'");
  worksheet.write_string(CELL("C1"), "Column#");
  worksheet.write_string(CELL("D1"), "Column[");
  worksheet.write_string(CELL("E1"), "Column]");

  const std::vector<xwpp::table_column_t> columns{
    {.header_ = "Column1", .total_string_ = "Total"                             },
    {.header_ = "Column'", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Column#", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Column[", .total_function_ = xwpp::table_total_functions_t::SUM},
    {.header_ = "Column]", .total_function_ = xwpp::table_total_functions_t::SUM}
  };
  const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

  worksheet.add_table(RANGE("B3:F9"), options);

  workbook.save("test_table23.xlsx");
}
