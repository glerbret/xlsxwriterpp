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

  worksheet.write_url(CELL("A1"), R"(external:..\foo.xlsx)");
  worksheet.write_url(CELL("A3"), R"(external:..\foo.xlsx#Sheet1!A1)");
  worksheet.write_url(CELL("A5"), R"(external:\\VBOXSVR\share\foo.xlsx#Sheet1!B2)", nullptr, "J:\\foo.xlsx#Sheet1!B2",
                      "");

  workbook.save("test_hyperlink09.xlsx");
}
