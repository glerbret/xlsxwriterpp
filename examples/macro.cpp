/*
 * An example of adding macros to a Xlsxwriter++ file using a VBA project
 * file extracted from an existing Excel .xlsm file.
 *
 * The vba_extract.py utility from the Xlsxwriter++ examples directory can be
 * used to extract the vbaProject.bin file.
 *
 * This example connects the macro to a button (the only Excel/VBA form object
 * supported by Xlsxwriter++) but that isn't a requirement for adding a macro
 * file to the workbook.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_column(COLS("A:A"), 30);

  // Add a macro file extracted from an Excel workbook.
  workbook.add_vba_project("vbaProject.bin");

  worksheet.write_string(2, 0, "Press the button to say hello.");

  xwpp::button_options_t options = {.caption_ = "Press Me", .macro_ = "say_hello", .width_ = 80, .height_ = 30};
  worksheet.insert_button(2, 1, options);

  // Note the xlsm extension of the filename
  workbook.save("macro.xlsm");
}
