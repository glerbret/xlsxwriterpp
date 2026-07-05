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

  worksheet.set_paper(9);
  worksheet.set_dpi(0, 200);

  worksheet.repeat_rows(0, 0);
  worksheet.repeat_columns(0, 0);

  worksheet.write_string(CELL("A1"), "Foo");

  workbook.save("test_repeat03.xlsx");
}
