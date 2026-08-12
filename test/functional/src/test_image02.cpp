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

  const xwpp::image_options_t options{.x_offset_ = 1, .y_offset_ = 2};

  worksheet.insert_image(CELL("D7"), "images/yellow.png", options);

  workbook.save("test_image02.xlsx");
}
