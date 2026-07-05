/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();

  worksheet1.embed_image(0, 0, "images/red.png");
  worksheet1.embed_image(2, 0, "images/blue.png");
  worksheet1.embed_image(4, 0, "images/yellow.png");

  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  worksheet2.embed_image(0, 0, "images/yellow.png");
  worksheet2.embed_image(2, 0, "images/red.png");
  worksheet2.embed_image(4, 0, "images/blue.png");

  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  worksheet3.embed_image(0, 0, "images/blue.png");
  worksheet3.embed_image(2, 0, "images/yellow.png");
  worksheet3.embed_image(4, 0, "images/red.png");

  workbook.save("test_embed_image13.xlsx");
}
