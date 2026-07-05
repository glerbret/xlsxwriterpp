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

  worksheet.print_area(RANGE("A1:XFD1048576"));

  worksheet.write_string(CELL("A1"), "Foo", nullptr);

  workbook.save("test_print_area07.xlsx");
}
