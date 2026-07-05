/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet01 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet02 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet03 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet04 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet05 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet06 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet07 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet08 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet09 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet10 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet11 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet12 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet13 = workbook.add_worksheet();

  worksheet01.write_string(CELL("A1"), "Foo");
  worksheet02.write_string(CELL("A1"), "Foo");
  worksheet03.write_string(CELL("A1"), "Foo");
  worksheet04.write_string(CELL("A1"), "Foo");
  worksheet05.write_string(CELL("A1"), "Foo");
  worksheet06.write_string(CELL("A1"), "Foo");
  worksheet07.write_string(CELL("A1"), "Foo");
  worksheet08.write_string(CELL("A1"), "Foo");
  worksheet09.write_string(CELL("A1"), "Foo");
  worksheet10.write_string(CELL("A1"), "Foo");
  worksheet11.write_string(CELL("A1"), "Foo");
  worksheet12.write_string(CELL("A1"), "Foo");
  worksheet13.write_string(CELL("A1"), "Foo");

  worksheet01.freeze_panes(CELL("A2"));
  worksheet02.freeze_panes(CELL("A3"));
  worksheet03.freeze_panes(CELL("B1"));
  worksheet04.freeze_panes(CELL("C1"));
  worksheet05.freeze_panes(CELL("B2"));
  worksheet06.freeze_panes(CELL("G4"));
  worksheet07.freeze_panes(3, 6, 3, 6, 1);
  worksheet08.split_panes(15, 0);
  worksheet09.split_panes(30, 0);
  worksheet10.split_panes(0, 8.46);
  worksheet11.split_panes(0, 17.57);
  worksheet12.split_panes(15, 8.46);
  worksheet13.split_panes(45, 54.14);

  workbook.save("test_panes01.xlsx");
}
