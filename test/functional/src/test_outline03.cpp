/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outline Columns");

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  const xwpp::row_col_options_t options6{.hidden_ = false, .level_ = 1, .collapsed_ = false};

  worksheet.write(CELL("A1"), "Month");
  worksheet.write(CELL("B1"), "Jan");
  worksheet.write(CELL("C1"), "Feb");
  worksheet.write(CELL("D1"), "Mar");
  worksheet.write(CELL("E1"), "Apr");
  worksheet.write(CELL("F1"), "May");
  worksheet.write(CELL("G1"), "Jun");
  worksheet.write(CELL("H1"), "Total");

  worksheet.write(CELL("A2"), "North");
  worksheet.write(CELL("B2"), 50);
  worksheet.write(CELL("C2"), 20);
  worksheet.write(CELL("D2"), 15);
  worksheet.write(CELL("E2"), 25);
  worksheet.write(CELL("F2"), 65);
  worksheet.write(CELL("G2"), 80);
  worksheet.write_formula(CELL("H2"), "=SUM(B2:G2)", nullptr, 255);

  worksheet.write(CELL("A3"), "South");
  worksheet.write(CELL("B3"), 10);
  worksheet.write(CELL("C3"), 20);
  worksheet.write(CELL("D3"), 30);
  worksheet.write(CELL("E3"), 50);
  worksheet.write(CELL("F3"), 50);
  worksheet.write(CELL("G3"), 50);
  worksheet.write_formula(CELL("H3"), "=SUM(B3:G3)", nullptr, 210);

  worksheet.write(CELL("A4"), "East");
  worksheet.write(CELL("B4"), 45);
  worksheet.write(CELL("C4"), 75);
  worksheet.write(CELL("D4"), 50);
  worksheet.write(CELL("E4"), 15);
  worksheet.write(CELL("F4"), 75);
  worksheet.write(CELL("G4"), 100);
  worksheet.write_formula(CELL("H4"), "=SUM(B4:G4)", nullptr, 360);

  worksheet.write(CELL("A5"), "West");
  worksheet.write(CELL("B5"), 15);
  worksheet.write(CELL("C5"), 15);
  worksheet.write(CELL("D5"), 55);
  worksheet.write(CELL("E5"), 35);
  worksheet.write(CELL("F5"), 20);
  worksheet.write(CELL("G5"), 50);
  worksheet.write_formula(CELL("H5"), "=SUM(B5:G5)", nullptr, 190);

  worksheet.write_formula(CELL("H6"), "=SUM(H2:H5)", bold, 1015);

  worksheet.set_row(0, bold);

  worksheet.set_column(COLS("A:A"), 10, bold);
  worksheet.set_column(COLS("H:H"), 10);
  worksheet.set_column(COLS("B:G"), 6, nullptr, options6);

  workbook.save("test_outline03.xlsx");
}
