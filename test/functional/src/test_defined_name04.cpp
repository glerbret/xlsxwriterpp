/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  /* xwpp::worksheet_t& worksheet = */ workbook.add_worksheet();

  workbook.define_name("\\__", "=Sheet1!$A$1");
  workbook.define_name("a3f6", "=Sheet1!$A$2");
  workbook.define_name("afoo.bar", "=Sheet1!$A$3");
  workbook.define_name("étude", "=Sheet1!$A$4");
  workbook.define_name("eésumé", "=Sheet1!$A$5");
  workbook.define_name("a", "=Sheet1!$A$6");

  workbook.save("test_defined_name04.xlsx");
}
