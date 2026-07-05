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

  workbook.add_vba_project("images/vbaProject03.bin");
  workbook.set_vba_name("MyWorkbook");
  worksheet.set_vba_name("MySheet1");

  worksheet.write_number(CELL("A1"), 123);

  workbook.save("test_macro02.xlsm");
}
