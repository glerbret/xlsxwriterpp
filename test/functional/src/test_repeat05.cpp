/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  worksheet1.set_paper(9);
  worksheet1.set_dpi(0, 200);

  worksheet3.set_paper(9);
  worksheet3.set_dpi(0, 200);

  worksheet1.repeat_rows(0, 0);
  worksheet3.repeat_rows(2, 3);
  worksheet3.repeat_columns(1, 5);

  worksheet1.write_string(CELL("A1"), "Foo");

  workbook.save("test_repeat05.xlsx");
}
