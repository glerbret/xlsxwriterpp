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

  xwpp::format_t* default_url_format = workbook.get_default_url_format();

  worksheet.write_url(CELL("A1"), "http://www.perl.org/", default_url_format);

  workbook.save("test_hyperlink78.xlsx");
}
