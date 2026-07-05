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
  /* xwpp::worksheet_t& worksheet2 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet3 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet4 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet5 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet6 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet7 =*/workbook.add_worksheet();
  xwpp::worksheet_t& worksheet8 = workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet9 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet10 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet11 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet12 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet13 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet14 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet15 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet16 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet17 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet18 =*/workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet19 =*/workbook.add_worksheet();
  xwpp::worksheet_t& worksheet20 = workbook.add_worksheet();

  worksheet8.set_first_sheet();
  worksheet20.activate();

  workbook.save("test_firstsheet01.xlsx");
}
