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

  const xwpp::button_options_t options{.macro = "my_macro", .width = 128, .height = 30};
  worksheet.insert_button("C2", options);

  workbook.save("test_button06.xlsx");
}
