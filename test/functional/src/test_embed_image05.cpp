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

  worksheet.write_dynamic_array_formula(0, 0, 2, 0, "=LEN(B1:B3)");

  worksheet.embed_image(8, 4, "images/red.png");

  workbook.save("test_embed_image05.xlsx");
}
