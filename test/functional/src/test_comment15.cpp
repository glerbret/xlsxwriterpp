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

  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bold();

  worksheet.write_string(CELL("A1"), "Foo", format1);
  worksheet.write_comment(CELL("B2"), "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment15.xlsx");
}
