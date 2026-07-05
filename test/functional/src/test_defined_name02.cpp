/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  /* xwpp::worksheet_t& worksheet = */ workbook.add_worksheet("sheet One");

  workbook.define_name("Sales", "='sheet One'!$G$1:$H$10");

  workbook.save("test_defined_name02.xlsx");
}
