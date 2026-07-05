/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet("Start");
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet("A & B");

  workbook.unset_default_url_format();

  worksheet1.write_url(CELL("A1"), "internal:'A & B'!A1", nullptr, "Jump to A & B", "");

  workbook.save("test_escapes05.xlsx");
}
