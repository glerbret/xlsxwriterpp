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
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet("Data Sheet");
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  xwpp::format_t* format = workbook.add_format();
  format->set_bold();

  worksheet1.write(0, 0, "Foo");
  worksheet1.write(1, 0, 123);

  worksheet3.write(1, 1, "Foo");
  worksheet3.write(2, 1, "Bar", format);
  worksheet3.write(3, 2, 234);

  workbook.save("test_simple02.xlsx");
}
