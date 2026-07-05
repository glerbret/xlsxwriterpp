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
  xwpp::format_t* format1      = workbook.add_format();

  format1->set_bold();

  worksheet.write_string(CELL("A1"), "Test", format1);
  worksheet.write_url(CELL("A3"), "http://www.python.org/");

  workbook.save("test_hyperlink31.xlsx");
}
