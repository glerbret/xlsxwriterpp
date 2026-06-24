/*
 * Example of how to use the Xlsxwriter++ library to write simple
 * array formulas.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some data for the formulas.
  worksheet.write_number(0, 1, 500);
  worksheet.write_number(1, 1, 10);
  worksheet.write_number(4, 1, 1);
  worksheet.write_number(5, 1, 2);
  worksheet.write_number(6, 1, 3);

  worksheet.write_number(0, 2, 300);
  worksheet.write_number(1, 2, 15);
  worksheet.write_number(4, 2, 20234);
  worksheet.write_number(5, 2, 21003);
  worksheet.write_number(6, 2, 10000);

  // Write an array formula that returns a single value.
  worksheet.write_array_formula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}");

  // Similar to above but using the RANGE macro.
  worksheet.write_array_formula(RANGE("A2:A2"), "{=SUM(B1:C1*B2:C2)}");

  // Write an array formula that returns a range of values.
  worksheet.write_array_formula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}");

  workbook.save("array_formula.xlsx");

}
