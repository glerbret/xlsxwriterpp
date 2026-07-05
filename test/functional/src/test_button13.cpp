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

  xwpp::button_options_t options = {.caption_ = "Hello", .macro_ = "say_hello"};
  workbook.set_vba_name("ThisWorkbook");
  worksheet.set_vba_name("Sheet1");
  worksheet.insert_button(CELL("C2"), options);
  workbook.add_vba_project("images/vbaProject02.bin");

  workbook.save("test_button13.xlsm");
}
