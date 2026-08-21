/*
 * An example of how to hide rows and columns using the Xlsxwriter++
 * library.
 *
 * In order to hide rows without setting each one, (of approximately 1 million
 * rows), Excel uses an optimization to hide all rows that don't have data. In
 * Xlsxwriter++ we replicate that using the worksheet_set_default_row()
 * function.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some data.
  worksheet.write_string(0, 3, "Some hidden columns.");
  worksheet.write_string(7, 0, "Some hidden rows.");

  // Hide all rows without data.
  worksheet.set_default_row(15, true);

  // Set the height of empty rows that we want to display even if it is
  // the default height.
  for(xwpp::row_num_t row_num = 1; row_num <= 6; row_num++)
  {
    worksheet.set_row(row_num, 15);
  }

  // Columns can be hidden explicitly. This doesn't increase the file size.
  xwpp::row_col_options_t options = {.hidden_ = true};
  worksheet.set_column(COLS("G:XFD"), 8.43, nullptr, options);

  workbook.save("hide_row_col.xlsx");
}
