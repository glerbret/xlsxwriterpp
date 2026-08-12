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

  const xwpp::image_options_t options{.x_offset_ = 13, .y_offset_ = 2};

  worksheet.set_row(1, 4.5);
  worksheet.set_row(2, 35.25);
  worksheet.set_column(2, 4, 3.29);
  worksheet.set_column(5, 5, 10.71);

  worksheet.insert_image(CELL("C2"), "images/logo.png", options);

  workbook.save("test_image15.xlsx");
}
