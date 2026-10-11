/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Foo");

  workbook.add_signed_vba_project("images/vbaProject05.bin", "images/vbaProject05Signature.bin");

  worksheet.set_column("A:A", 30);
  worksheet.write(2, 0, "Press the button to say hello.");

  const xwpp::button_options_t options{.caption = "Press Me", .macro = "say_hello", .width = 80, .height = 30};

  worksheet.insert_button(2, 1, options);

  workbook.save("test_macro04.xlsm");
}
