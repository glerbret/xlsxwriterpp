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

  const xwpp::image_options_t options{.x_scale_ = 0.5, .y_scale_ = 0.5};

  worksheet.insert_image(CELL("B3"), "images/grey.png", options);

  workbook.save("test_image08.xlsx");
}
