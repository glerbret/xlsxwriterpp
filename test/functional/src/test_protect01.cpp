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

  const xwpp::format_t* unlocked = workbook.format_builder().unlocked().build();
  const xwpp::format_t* hidden   = workbook.format_builder().unlocked().hidden().build();

  worksheet.write("A1", 1);
  worksheet.write("A2", 2, unlocked);
  worksheet.write("A3", 3, hidden);

  workbook.save("test_protect01.xlsx");
}
