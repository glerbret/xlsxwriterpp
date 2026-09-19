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

  worksheet.set_column("B:K", 10.288);

  worksheet.write("A1", "Column1");
  worksheet.write("B1", "Column2");
  worksheet.write("C1", "Column3");
  worksheet.write("D1", "Column4");
  worksheet.write("E1", "Column5");
  worksheet.write("F1", "Column6");
  worksheet.write("G1", "Column7");
  worksheet.write("H1", "Column8");
  worksheet.write("I1", "Column9");
  worksheet.write("J1", "Column10");
  worksheet.write("K1", "Total");

  worksheet.write(3, 1, 0);
  worksheet.write(3, 2, 0);
  worksheet.write(3, 3, 0);
  worksheet.write(3, 6, 4);
  worksheet.write(3, 7, 0);
  worksheet.write(3, 8, 1);
  worksheet.write(3, 9, 0);
  worksheet.write(3, 10, 0);

  worksheet.write(4, 1, 0);
  worksheet.write(4, 2, 0);
  worksheet.write(4, 3, 0);
  worksheet.write(4, 6, 5);
  worksheet.write(4, 7, 0);
  worksheet.write(4, 8, 2);
  worksheet.write(4, 9, 0);
  worksheet.write(4, 10, 0);

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
  worksheet.add_table("B3:K6", options);

  workbook.save("test_table17.xlsx");
}
