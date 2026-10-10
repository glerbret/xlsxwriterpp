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

  const xwpp::format_t* bold = workbook.format_builder().bold().build();

  worksheet.write("A1", "Foo", bold);
  worksheet.write_comment("B2", "Some text");

  worksheet.set_comments_author("John");

  workbook.save("test_comment15.xlsx");
}
