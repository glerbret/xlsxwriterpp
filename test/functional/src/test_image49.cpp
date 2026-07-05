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
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  worksheet1.insert_image(CELL("A1"), "images/blue.png");
  worksheet1.insert_image(CELL("B3"), "images/red.jpg");
  worksheet1.insert_image(CELL("D5"), "images/yellow.jpg");
  worksheet1.insert_image(CELL("F9"), "images/grey.png");

  worksheet2.insert_image(CELL("A1"), "images/blue.png");
  worksheet2.insert_image(CELL("B3"), "images/red.jpg");
  worksheet2.insert_image(CELL("D5"), "images/yellow.jpg");
  worksheet2.insert_image(CELL("F9"), "images/grey.png");

  worksheet3.insert_image(CELL("A1"), "images/blue.png");
  worksheet3.insert_image(CELL("B3"), "images/red.jpg");
  worksheet3.insert_image(CELL("D5"), "images/yellow.jpg");
  worksheet3.insert_image(CELL("F9"), "images/grey.png");

  workbook.save("test_image49.xlsx");
}
