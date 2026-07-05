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
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet("Data Sheet");
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  worksheet1.write_string(CELL("A1"), "Foo");
  worksheet1.write_number(CELL("A2"), 123);

  worksheet3.write_string(CELL("B2"), "Foo");
  worksheet3.write_string(CELL("B3"), "Bar", bold);
  worksheet3.write_number(CELL("C4"), 234);

  // Ensure the active worksheet is overwritten, below.
  worksheet2.activate();

  worksheet2.select();
  worksheet3.select();
  worksheet3.activate();

  workbook.save("test_simple03.xlsx");
}
