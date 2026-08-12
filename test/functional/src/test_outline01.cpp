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

  worksheet.set_column(COLS("A:A"), 20);

  worksheet.set_row(1, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(2, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(3, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(4, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(5, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

  worksheet.set_row(6, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(7, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(8, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(9, xwpp::DEF_ROW_HEIGHT, nullptr, options1);
  worksheet.set_row(10, xwpp::DEF_ROW_HEIGHT, nullptr, options2);

  worksheet.write_string(CELL("A1"), "Region", bold);
  worksheet.write_string(CELL("A2"), "North");
  worksheet.write_string(CELL("A3"), "North");
  worksheet.write_string(CELL("A4"), "North");
  worksheet.write_string(CELL("A5"), "North");
  worksheet.write_string(CELL("A6"), "North Total", bold);

  worksheet.write_string(CELL("B1"), "Sales", bold);
  worksheet.write_number(CELL("B2"), 1000);
  worksheet.write_number(CELL("B3"), 1200);
  worksheet.write_number(CELL("B4"), 900);
  worksheet.write_number(CELL("B5"), 1200);
  worksheet.write_formula_num(CELL("B6"), "=SUBTOTAL(9,B2:B5)", bold, 4300);

  worksheet.write_string(CELL("A7"), "South");
  worksheet.write_string(CELL("A8"), "South");
  worksheet.write_string(CELL("A9"), "South");
  worksheet.write_string(CELL("A10"), "South");
  worksheet.write_string(CELL("A11"), "South Total", bold);

  worksheet.write_number(CELL("B7"), 400);
  worksheet.write_number(CELL("B8"), 600);
  worksheet.write_number(CELL("B9"), 500);
  worksheet.write_number(CELL("B10"), 600);
  worksheet.write_formula_num(CELL("B11"), "=SUBTOTAL(9,B7:B10)", bold, 2100);

  worksheet.write_string(CELL("A12"), "Grand Total", bold);
  worksheet.write_formula_num(CELL("B12"), "=SUBTOTAL(9,B2:B10)", bold, 6400);

  workbook.save("test_outline01.xlsx");
}
