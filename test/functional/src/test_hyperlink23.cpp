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

  worksheet.write_url(CELL("A1"), "https://en.wikipedia.org/wiki/Microsoft_Excel#Data_storage_and_communication",
                      nullptr, "Display text", "");

  workbook.save("test_hyperlink23.xlsx");
}
