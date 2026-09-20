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

  worksheet.write_url("A1", "external:C:\\Temp\\foo.xlsx");
  worksheet.write_url("A3", "external:C:\\Temp\\foo.xlsx#Sheet1!A1");
  worksheet.write_url("A5", "external:C:\\Temp\\foo.xlsx#Sheet1!A1", "External", "Tip");

  workbook.save("test_hyperlink06.xlsx");
}
