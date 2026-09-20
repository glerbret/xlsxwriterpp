/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet("Collapsed Rows");

  xwpp::format_t* bold = workbook.add_format();
  bold->set_bold();

  const xwpp::row_col_options_t options3{.hidden_ = true, .level_ = 2, .collapsed_ = false};
  const xwpp::row_col_options_t options4{.hidden_ = true, .level_ = 1, .collapsed_ = false};
  const xwpp::row_col_options_t options5{.hidden_ = false, .level_ = 0, .collapsed_ = true};

  worksheet.set_column("A:A", 20);
  worksheet.set_selection("A14:A14");

  worksheet.set_row(1, options3);
  worksheet.set_row(2, options3);
  worksheet.set_row(3, options3);
  worksheet.set_row(4, options3);
  worksheet.set_row(5, options4);

  worksheet.set_row(6, options3);
  worksheet.set_row(7, options3);
  worksheet.set_row(8, options3);
  worksheet.set_row(9, options3);
  worksheet.set_row(10, options4);
  worksheet.set_row(11, options5);

  worksheet.write("A1", "Region", bold);
  worksheet.write("A2", "North");
  worksheet.write("A3", "North");
  worksheet.write("A4", "North");
  worksheet.write("A5", "North");
  worksheet.write("A6", "North Total", bold);

  worksheet.write("B1", "Sales", bold);
  worksheet.write("B2", 1000);
  worksheet.write("B3", 1200);
  worksheet.write("B4", 900);
  worksheet.write("B5", 1200);
  worksheet.write_formula("B6", "=SUBTOTAL(9,B2:B5)", bold, 4300);

  worksheet.write("A7", "South");
  worksheet.write("A8", "South");
  worksheet.write("A9", "South");
  worksheet.write("A10", "South");
  worksheet.write("A11", "South Total", bold);

  worksheet.write("B7", 400);
  worksheet.write("B8", 600);
  worksheet.write("B9", 500);
  worksheet.write("B10", 600);
  worksheet.write_formula("B11", "=SUBTOTAL(9,B7:B10)", bold, 2100);

  worksheet.write("A12", "Grand Total", bold);
  worksheet.write_formula("B12", "=SUBTOTAL(9,B2:B10)", bold, 6400);

  workbook.save("test_outline02.xlsx");
}
