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

  xwpp::image_options_t image_options = {.description_ = "Some alt text"};

  worksheet.embed_image(0, 0, "images/red.png", image_options);

  workbook.save("test_embed_image08.xlsx");
}
