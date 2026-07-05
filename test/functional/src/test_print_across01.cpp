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

  worksheet.print_across();
  worksheet.set_paper(9);
  worksheet.set_dpi(0, 200);

  worksheet.write_string(0, 0, "Foo");

  workbook.save("test_print_across01.xlsx");
}
