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
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet3 = */ workbook.add_worksheet("Sheet 3");

  worksheet1.set_paper(9);
  worksheet1.set_dpi(0, 200);

  worksheet1.print_area(RANGE("A1:E6"));
  worksheet1.autofilter(RANGE("F1:G1"));
  worksheet1.write_string(CELL("G1"), "Filter");
  worksheet1.write_string(CELL("F1"), "Auto");
  worksheet1.fit_to_pages(2, 2);

  workbook.define_name("'Sheet 3'!Bar", "='Sheet 3'!$A$1");
  workbook.define_name("Abc", "=Sheet1!$A$1");
  workbook.define_name("Baz", "=0.98");
  workbook.define_name("Sheet1!Bar", "=Sheet1!$A$1");
  workbook.define_name("Sheet2!Bar", "=Sheet2!$A$1");
  workbook.define_name("Sheet2!aaa", "=Sheet2!$A$1");
  workbook.define_name("_Egg", "=Sheet1!$A$1");
  workbook.define_name("_Fog", "=Sheet1!$A$1");

  workbook.save("test_defined_name01.xlsx");
}
