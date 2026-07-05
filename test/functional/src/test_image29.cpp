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

  xwpp::image_options_t options = {.x_offset_ = -210, .y_offset_ = 1};

  worksheet.insert_image(0, 10, "images/red_208.png", options);

  workbook.save("test_image29.xlsx");
}
