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

  worksheet.set_column(COLS("B:B"), 5);

  const xwpp::image_options_t image_options{.x_offset_ = 232};
  worksheet.insert_image(CELL("A9"), "images/red.png", image_options);

  workbook.save("test_object_position15.xlsx");
}
