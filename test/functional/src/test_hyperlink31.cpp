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
  const xwpp::format_t* bold   = workbook.format_builder().bold().build();

  worksheet.write("A1", "Test", bold);
  worksheet.write_url("A3", "http://www.python.org/");

  workbook.save("test_hyperlink31.xlsx");
}
