/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_column(COLS("C:F"), 10.288);

  worksheet.write_string(CELL("A1"), "Column1");
  worksheet.write_string(CELL("B1"), "Column2");
  worksheet.write_string(CELL("C1"), "Column3");
  worksheet.write_string(CELL("D1"), "Column4");
  worksheet.write_string(CELL("E1"), "Total");

  std::vector<xwpp::table_column_t> columns = {
    {.total_string_ = "Total"}, {}, {}, {.total_function_ = xwpp::table_total_functions_t::COUNT}};

  xwpp::table_options_t options = {.total_row_ = true, .columns_ = columns};

  worksheet.add_table(RANGE("C3:F14"), options);

  workbook.save("test_table08.xlsx");
}
