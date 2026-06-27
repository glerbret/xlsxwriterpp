/*
 * An example of writing cell comments to a worksheet using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.write_string(0, 0, "Hello");
  worksheet.write_comment(0, 0, "This is a comment");

  workbook.save("comments1.xlsx");
}
