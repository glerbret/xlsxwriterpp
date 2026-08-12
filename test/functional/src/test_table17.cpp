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

  worksheet.set_column(COLS("B:K"), 10.288);

  worksheet.write_string(CELL("A1"), "Column1");
  worksheet.write_string(CELL("B1"), "Column2");
  worksheet.write_string(CELL("C1"), "Column3");
  worksheet.write_string(CELL("D1"), "Column4");
  worksheet.write_string(CELL("E1"), "Column5");
  worksheet.write_string(CELL("F1"), "Column6");
  worksheet.write_string(CELL("G1"), "Column7");
  worksheet.write_string(CELL("H1"), "Column8");
  worksheet.write_string(CELL("I1"), "Column9");
  worksheet.write_string(CELL("J1"), "Column10");
  worksheet.write_string(CELL("K1"), "Total");

  worksheet.write_number(3, 1, 0);
  worksheet.write_number(3, 2, 0);
  worksheet.write_number(3, 3, 0);
  worksheet.write_number(3, 6, 4);
  worksheet.write_number(3, 7, 0);
  worksheet.write_number(3, 8, 1);
  worksheet.write_number(3, 9, 0);
  worksheet.write_number(3, 10, 0);

  worksheet.write_number(4, 1, 0);
  worksheet.write_number(4, 2, 0);
  worksheet.write_number(4, 3, 0);
  worksheet.write_number(4, 6, 5);
  worksheet.write_number(4, 7, 0);
  worksheet.write_number(4, 8, 2);
  worksheet.write_number(4, 9, 0);
  worksheet.write_number(4, 10, 0);

  const std::vector<xwpp::table_column_t> columns{
    {.total_string_ = "Total"},
    {},
    {.total_function_ = xwpp::table_total_functions_t::AVERAGE},
    {.total_function_ = xwpp::table_total_functions_t::COUNT},
    {.total_function_ = xwpp::table_total_functions_t::COUNT_NUMS},
    {.total_function_ = xwpp::table_total_functions_t::MAX, .total_value_ = 5},
    {.total_function_ = xwpp::table_total_functions_t::MIN},
    {.total_function_ = xwpp::table_total_functions_t::SUM, .total_value_ = 3},
    {.total_function_ = xwpp::table_total_functions_t::STD_DEV},
    {.total_function_ = xwpp::table_total_functions_t::VAR}
  };

  const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

  worksheet.add_table(RANGE("B3:K6"), options);

  workbook.save("test_table17.xlsx");
}
