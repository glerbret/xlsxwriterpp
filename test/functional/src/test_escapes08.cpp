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

  workbook.unset_default_url_format();

  // Test an already escaped string.
  worksheet.write_url(CELL("A1"), "http://example.com/%5b0%5d", nullptr, "http://example.com/[0]", "");

  workbook.save("test_escapes08.xlsx");
}
