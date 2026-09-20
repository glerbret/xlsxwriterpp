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

  worksheet.set_column("C:F", 10.288);

  worksheet.write("A1", "Column1");
  worksheet.write("B1", "Column2");
  worksheet.write("C1", "Column3");
  worksheet.write("D1", "Column4");
  worksheet.write("E1", "Total");

  const std::vector<xwpp::table_column_t> columns{
    {.total_string_ = "Total"}, {}, {}, {.total_function_ = xwpp::table_total_functions_t::COUNT}};

  const xwpp::table_options_t options{.total_row_ = true, .columns_ = columns};

  worksheet.add_table("C3:F14", options);

  workbook.save("test_table08.xlsx");
}
