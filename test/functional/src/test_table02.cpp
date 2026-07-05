/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  worksheet1.set_column(COLS("B:J"), 10.288);
  worksheet2.set_column(COLS("C:L"), 10.288);

  worksheet1.add_table(RANGE("B3:E11"));
  worksheet1.add_table(RANGE("G10:J16"));
  worksheet1.add_table(RANGE("C18:F25"));

  worksheet2.add_table(RANGE("I4:L11"));
  worksheet2.add_table(RANGE("C16:H23"));

  workbook.save("test_table02.xlsx");
}
