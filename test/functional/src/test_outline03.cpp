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

  worksheet.write_string(CELL("A1"), "Month");
  worksheet.write_string(CELL("B1"), "Jan");
  worksheet.write_string(CELL("C1"), "Feb");
  worksheet.write_string(CELL("D1"), "Mar");
  worksheet.write_string(CELL("E1"), "Apr");
  worksheet.write_string(CELL("F1"), "May");
  worksheet.write_string(CELL("G1"), "Jun");
  worksheet.write_string(CELL("H1"), "Total");

  worksheet.write_string(CELL("A2"), "North");
  worksheet.write_number(CELL("B2"), 50);
  worksheet.write_number(CELL("C2"), 20);
  worksheet.write_number(CELL("D2"), 15);
  worksheet.write_number(CELL("E2"), 25);
  worksheet.write_number(CELL("F2"), 65);
  worksheet.write_number(CELL("G2"), 80);
  worksheet.write_formula_num(CELL("H2"), "=SUM(B2:G2)", nullptr, 255);

  worksheet.write_string(CELL("A3"), "South");
  worksheet.write_number(CELL("B3"), 10);
  worksheet.write_number(CELL("C3"), 20);
  worksheet.write_number(CELL("D3"), 30);
  worksheet.write_number(CELL("E3"), 50);
  worksheet.write_number(CELL("F3"), 50);
  worksheet.write_number(CELL("G3"), 50);
  worksheet.write_formula_num(CELL("H3"), "=SUM(B3:G3)", nullptr, 210);

  worksheet.write_string(CELL("A4"), "East");
  worksheet.write_number(CELL("B4"), 45);
  worksheet.write_number(CELL("C4"), 75);
  worksheet.write_number(CELL("D4"), 50);
  worksheet.write_number(CELL("E4"), 15);
  worksheet.write_number(CELL("F4"), 75);
  worksheet.write_number(CELL("G4"), 100);
  worksheet.write_formula_num(CELL("H4"), "=SUM(B4:G4)", nullptr, 360);

  worksheet.write_string(CELL("A5"), "West");
  worksheet.write_number(CELL("B5"), 15);
  worksheet.write_number(CELL("C5"), 15);
  worksheet.write_number(CELL("D5"), 55);
  worksheet.write_number(CELL("E5"), 35);
  worksheet.write_number(CELL("F5"), 20);
  worksheet.write_number(CELL("G5"), 50);
  worksheet.write_formula_num(CELL("H5"), "=SUM(B5:G5)", nullptr, 190);

  worksheet.write_formula_num(CELL("H6"), "=SUM(H2:H5)", bold, 1015);

  worksheet.set_row(0, xwpp::DEF_ROW_HEIGHT, bold, std::nullopt);

  worksheet.set_column(COLS("A:A"), 10, bold, std::nullopt);
  worksheet.set_column(COLS("H:H"), 10);
  worksheet.set_column(COLS("B:G"), 6, nullptr, options6);

  workbook.save("test_outline03.xlsx");
}
