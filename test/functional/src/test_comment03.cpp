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

  worksheet.write("A1", "Foo");
  worksheet.write_comment("A1", "Some text");
  worksheet.write_comment("XFD1048576", "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment03.xlsx");
}
