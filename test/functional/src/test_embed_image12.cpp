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

  const xwpp::format_t* format = workbook.format_builder().bg_color(xwpp::color_t{0xFFFF00}).build();

  const xwpp::image_options_t image_options{.cell_format_ = format};

  worksheet.embed_image(0, 0, "images/red.png", image_options);

  workbook.save("test_embed_image12.xlsx");
}
