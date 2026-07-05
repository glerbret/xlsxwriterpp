/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.write_string(CELL("A1"), "Foo", NULL);
  worksheet.write_comment(CELL("A1"), "Some text");
  worksheet.write_comment(CELL("XFD1048576"), "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment03.xlsx");
}
