/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Outlined Rows");

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  const xwpp::row_col_options_t options1{.hidden_ = false, .level_ = 2, .collapsed_ = false};
  const xwpp::row_col_options_t options2{.hidden_ = false, .level_ = 1, .collapsed_ = false};

  worksheet.set_column("A:A", 20);

  worksheet.set_row(1, options1);
  worksheet.set_row(2, options1);
  worksheet.set_row(3, options1);
  worksheet.set_row(4, options1);
  worksheet.set_row(5, options2);

  worksheet.set_row(6, options1);
  worksheet.set_row(7, options1);
  worksheet.set_row(8, options1);
  worksheet.set_row(9, options1);
  worksheet.set_row(10, options2);

  worksheet.write(CELL("A1"), "Region", bold);
  worksheet.write(CELL("A2"), "North");
  worksheet.write(CELL("A3"), "North");
  worksheet.write(CELL("A4"), "North");
  worksheet.write(CELL("A5"), "North");
  worksheet.write(CELL("A6"), "North Total", bold);

  worksheet.write(CELL("B1"), "Sales", bold);
  worksheet.write(CELL("B2"), 1000);
  worksheet.write(CELL("B3"), 1200);
  worksheet.write(CELL("B4"), 900);
  worksheet.write(CELL("B5"), 1200);
  worksheet.write_formula(CELL("B6"), "=SUBTOTAL(9,B2:B5)", bold, 4300);

  worksheet.write(CELL("A7"), "South");
  worksheet.write(CELL("A8"), "South");
  worksheet.write(CELL("A9"), "South");
  worksheet.write(CELL("A10"), "South");
  worksheet.write(CELL("A11"), "South Total", bold);

  worksheet.write(CELL("B7"), 400);
  worksheet.write(CELL("B8"), 600);
  worksheet.write(CELL("B9"), 500);
  worksheet.write(CELL("B10"), 600);
  worksheet.write_formula(CELL("B11"), "=SUBTOTAL(9,B7:B10)", bold, 2100);

  worksheet.write(CELL("A12"), "Grand Total", bold);
  worksheet.write_formula(CELL("B12"), "=SUBTOTAL(9,B2:B10)", bold, 6400);

  workbook.save("test_outline01.xlsx");
}
