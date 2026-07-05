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

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  worksheet.write_boolean(CELL("A1"), true, bold);
  worksheet.write_boolean(CELL("A2"), false, italic);

  workbook.save("test_types08.xlsx");
}
