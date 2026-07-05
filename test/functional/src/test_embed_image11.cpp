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

  worksheet.set_column(0, 0, 11);
  worksheet.set_row(0, 72);

  worksheet.embed_image(0, 0, "images/red.png");

  workbook.save("test_embed_image11.xlsx");
}
