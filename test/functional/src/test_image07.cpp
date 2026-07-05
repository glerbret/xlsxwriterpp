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
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  worksheet1.insert_image(CELL("E9"), "images/red.png");
  worksheet2.insert_image(CELL("E9"), "images/yellow.png");

  workbook.save("test_image07.xlsx");
}
