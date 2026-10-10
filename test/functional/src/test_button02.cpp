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

  const xwpp::button_options_t options{.caption = "my text", .x_offset = 4, .y_offset = 3};
  worksheet.insert_button("B4", options);

  workbook.save("test_button02.xlsx");
}
