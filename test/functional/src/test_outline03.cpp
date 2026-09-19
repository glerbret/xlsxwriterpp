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

  worksheet.write("A1", "Month");
  worksheet.write("B1", "Jan");
  worksheet.write("C1", "Feb");
  worksheet.write("D1", "Mar");
  worksheet.write("E1", "Apr");
  worksheet.write("F1", "May");
  worksheet.write("G1", "Jun");
  worksheet.write("H1", "Total");

  worksheet.write("A2", "North");
  worksheet.write("B2", 50);
  worksheet.write("C2", 20);
  worksheet.write("D2", 15);
  worksheet.write("E2", 25);
  worksheet.write("F2", 65);
  worksheet.write("G2", 80);
  worksheet.write_formula("H2", "=SUM(B2:G2)", nullptr, 255);

  worksheet.write("A3", "South");
  worksheet.write("B3", 10);
  worksheet.write("C3", 20);
  worksheet.write("D3", 30);
  worksheet.write("E3", 50);
  worksheet.write("F3", 50);
  worksheet.write("G3", 50);
  worksheet.write_formula("H3", "=SUM(B3:G3)", nullptr, 210);

  worksheet.write("A4", "East");
  worksheet.write("B4", 45);
  worksheet.write("C4", 75);
  worksheet.write("D4", 50);
  worksheet.write("E4", 15);
  worksheet.write("F4", 75);
  worksheet.write("G4", 100);
  worksheet.write_formula("H4", "=SUM(B4:G4)", nullptr, 360);

  worksheet.write("A5", "West");
  worksheet.write("B5", 15);
  worksheet.write("C5", 15);
  worksheet.write("D5", 55);
  worksheet.write("E5", 35);
  worksheet.write("F5", 20);
  worksheet.write("G5", 50);
  worksheet.write_formula("H5", "=SUM(B5:G5)", nullptr, 190);

  worksheet.write_formula("H6", "=SUM(H2:H5)", bold, 1015);

  worksheet.set_row(0, bold);

  worksheet.set_column("A:A", 10, bold);
  worksheet.set_column("H:H", 10);
  worksheet.set_column("B:G", 6, nullptr, options6);

  workbook.save("test_outline03.xlsx");
}
