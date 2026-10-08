/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  workbook.unset_default_url_format();
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.write_url("B2", "external:./subdir/blank.xlsx");

  workbook.save("test_hyperlink16.xlsx");
}
