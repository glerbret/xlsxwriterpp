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

  worksheet.embed_image(0, 0, "images/red.png");
  worksheet.insert_image(8, 4, "images/red.png");

  workbook.save("test_embed_image06.xlsx");
}
