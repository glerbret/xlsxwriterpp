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

  xwpp::row_col_options_t options = {.hidden_ = true};

  worksheet.set_row_pixels(0, 1);
  worksheet.set_row_pixels(1, 2);
  worksheet.set_row_pixels(2, 3);
  worksheet.set_row_pixels(3, 4);

  worksheet.set_row_pixels(11, 12);
  worksheet.set_row_pixels(12, 13);
  worksheet.set_row_pixels(13, 14);
  worksheet.set_row_pixels(14, 15);

  worksheet.set_row_pixels(18, 19);
  worksheet.set_row_pixels(20, 21, nullptr, options);
  worksheet.set_row_pixels(21, 22);

  workbook.save("test_set_row04.xlsx");
}
