/*
 * An example of using the new Excel LAMBDA() function with the Xlsxwriter++
 * library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Note that the formula name is prefixed with "_xlfn." and that the
  // lambda function parameters are prefixed with "_xlpm.". These prefixes
  // won't show up in Excel.
  worksheet.write_dynamic_formula(CELL("A1"), "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)");

  // Create the lambda function as a defined name and write it as a dynamic formula.
  workbook.define_name("ToCelsius", "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))");

  worksheet.write_dynamic_formula(CELL("A2"), "=ToCelsius(212)");

  workbook.save("lambda.xlsx");
}
