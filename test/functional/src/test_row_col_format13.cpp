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

  const xwpp::row_col_options_t options{.hidden_ = true, .level_ = 0, .collapsed_ = false};

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  worksheet.set_column("B:D", 5);
  worksheet.set_column("F:F", 8, nullptr, options);
  worksheet.set_column("H:H", bold);
  worksheet.set_column("J:J", 2);
  worksheet.set_column("L:L", options);

  workbook.save("test_row_col_format13.xlsx");
}
