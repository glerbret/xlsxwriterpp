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

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  worksheet.set_row(1048575, 15, bold);

  worksheet.write_string(1048575, 0, "Bar");

  workbook.save("test_row_col_format18.xlsx");
}
