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

  xwpp::button_options_t options = {.macro_ = "my_macro", .width_ = 128, .height_ = 30};
  worksheet.insert_button(CELL("C2"), options);

  workbook.save("test_button06.xlsx");
}
