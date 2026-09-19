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

  worksheet.set_column("B:F", 10.288);

  worksheet.write(CELL("A1"), "Column1");
  worksheet.write(CELL("F1"), "Total");
  worksheet.write(CELL("B1"), "Column'");
  worksheet.write(CELL("C1"), "Column#");
  worksheet.write(CELL("D1"), "Column[");
  worksheet.write(CELL("E1"), "Column]");

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
