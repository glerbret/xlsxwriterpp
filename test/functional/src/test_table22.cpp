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

  worksheet.set_column("B:C", 10.288);

  worksheet.write(CELL("B2"), "apple", nullptr);
  worksheet.write(CELL("C2"), "pie", nullptr);
  worksheet.write(CELL("B3"), "pine", nullptr);
  worksheet.write(CELL("C3"), "tree", nullptr);

  const xwpp::table_options_t options{.no_header_row_ = true};

  worksheet.add_table(RANGE("B2:C3"), options);

  workbook.save("test_table22.xlsx");
}
