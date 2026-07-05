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

  xwpp::format_t* unlocked = workbook.add_format();
  unlocked->set_unlocked();

  xwpp::format_t* hidden = workbook.add_format();
  hidden->set_unlocked();
  hidden->set_hidden();

  worksheet.protect();

  worksheet.write_number(CELL("A1"), 1);
  worksheet.write_number(CELL("A2"), 2, unlocked);
  worksheet.write_number(CELL("A3"), 3, hidden);

  workbook.save("test_protect02.xlsx");
}
