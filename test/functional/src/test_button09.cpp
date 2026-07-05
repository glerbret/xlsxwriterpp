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

  worksheet2.insert_button(CELL("C2"));
  worksheet1.write_comment(CELL("A1"), "Foo");
  worksheet1.set_comments_author("John");

  workbook.save("test_button09.xlsx");
}
