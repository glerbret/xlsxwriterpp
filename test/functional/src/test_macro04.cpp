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

  worksheet.set_column(COLS("A:A"), 30);
  worksheet.write_string(2, 0, "Press the button to say hello.");

  xwpp::button_options_t options = {.caption_ = "Press Me", .macro_ = "say_hello", .width_ = 80, .height_ = 30};

  worksheet.insert_button(2, 1, options);

  workbook.save("test_macro04.xlsm");
}
