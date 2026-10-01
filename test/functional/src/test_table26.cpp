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

  worksheet.set_column("C:D", 10.288);
  worksheet.set_column("F:G", 10.288);

  // Add some valid tables.
  worksheet.add_table("C2:D3");

  const xwpp::table_options_t options1{.name_ = "Table2", .no_header_row_ = true};
  worksheet.add_table("F3:G3", options1);

  workbook.save("test_table26.xlsx");
}
