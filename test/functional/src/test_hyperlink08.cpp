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

  // Test with forward slashes instead of back slashes in test_hyperlink07.c.
  worksheet.write_url(CELL("A1"), "external://VBOXSVR/share/foo.xlsx", nullptr, "J:/foo.xlsx", "");
  worksheet.write_url(CELL("A3"), "external:foo.xlsx");

  workbook.save("test_hyperlink08.xlsx");
}
