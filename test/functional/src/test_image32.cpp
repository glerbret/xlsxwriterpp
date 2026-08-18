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

  const xwpp::image_options_t options{.x_offset_ = -100, .y_offset_ = -100};

  worksheet.insert_image(CELL("B2"), "images/red.png", options);

  workbook.save("test_image32.xlsx");
}
