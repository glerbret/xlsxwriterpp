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
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  worksheet1.write_string(CELL("A1"), "Foo");
  worksheet1.write_comment(CELL("B2"), "Some text");

  worksheet3.write_string(CELL("A1"), "Bar");
  worksheet3.write_comment(CELL("C7"), "More text");

  worksheet1.set_comments_author("John");
  worksheet3.set_comments_author("John");

  workbook.save("test_comment04.xlsx");
}
