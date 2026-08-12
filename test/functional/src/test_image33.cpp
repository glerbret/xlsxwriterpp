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

  const xwpp::image_options_t options{.x_offset_ = -2, .y_offset_ = -1};

  worksheet.set_column(3, 3, 3.86);
  worksheet.set_column(4, 4, 1.43);
  worksheet.set_row(7, 7.5);
  worksheet.set_row(8, 9.75);

  worksheet.insert_image(CELL("E9"), "images/red.png", options);

  workbook.save("test_image33.xlsx");
}
