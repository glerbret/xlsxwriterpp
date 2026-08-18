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

  const xwpp::image_options_t options{.x_offset_ = 8, .y_offset_ = 5};

  worksheet.insert_image(CELL("C2"), "images/logo.png", options);

  workbook.save("test_image11.xlsx");
}
