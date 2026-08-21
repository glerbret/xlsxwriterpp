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
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet5 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet6 = workbook.add_worksheet();

  worksheet1.set_selection(3, 2, 3, 2);     // 1. Cell C4.
  worksheet2.set_selection(3, 2, 6, 6);     // 2. Cells C4 to G7.
  worksheet3.set_selection(6, 6, 3, 2);     // 3. Cells G7 to C.4
  worksheet4.set_selection(RANGE("C4:C4")); // Same as 1.
  worksheet5.set_selection(RANGE("C4:G7")); // Same as 2.
  worksheet6.set_selection(RANGE("G7:C4")); // Same as 3.

  workbook.save("test_set_selection02.xlsx");
}
