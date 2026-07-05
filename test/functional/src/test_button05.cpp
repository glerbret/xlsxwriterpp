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

  xwpp::button_options_t options = {.macro_ = "my_macro", .x_scale_ = 2, .y_scale_ = 1.5};
  worksheet.insert_button(CELL("C2"), options);

  workbook.save("test_button05.xlsx");
}
