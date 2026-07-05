/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  /* xwpp::worksheet_t& worksheet1 =*/workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  /*xwpp::worksheet_t& worksheet3 =*/workbook.add_worksheet();

  worksheet2.hide();

  workbook.save("test_hide01.xlsx");
}
