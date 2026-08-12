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

  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bg_color(xwpp::color_t(0xFFFF00));

  const xwpp::image_options_t image_options{.cell_format_ = format1};

  worksheet.embed_image(0, 0, "images/red.png", image_options);

  workbook.save("test_embed_image12.xlsx");
}
