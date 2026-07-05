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

  worksheet.insert_image(CELL("B2"), "images/black_72.jpg");
  worksheet.insert_image(CELL("B8"), "images/black_96.jpg");
  worksheet.insert_image(CELL("B13"), "images/black_150.jpg");
  worksheet.insert_image(CELL("B17"), "images/black_300.jpg");

  workbook.save("test_image23.xlsx");
}
