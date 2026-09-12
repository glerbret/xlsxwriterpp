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

  xwpp::format_t* mixed = workbook.add_format();
  mixed->set_bold();
  mixed->set_italic();

  xwpp::format_t* italic = workbook.add_format();
  italic->set_italic();

  // Manually force the format index order for testing.
  workbook.set_default_xf_indices();

  worksheet.set_row(0, 15, bold);
  worksheet.set_column(0, 0, 8.43, italic);

  worksheet.write(0, 0, "Foo", mixed);
  worksheet.write(0, 1, "Foo");
  worksheet.write(1, 0, "Foo");
  worksheet.write(1, 1, "Foo");

  workbook.save("test_row_col_format08.xlsx");
}
