/*
 * Example of how to create defined names using Xlsxwriter++. This method is
 * used to define a user friendly name to represent a value, a single cell or
 * a range of cells in a workbook.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;

  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();

  // Define some global/workbook names.
  workbook.define_name("Exchange_rate", "=0.96");
  workbook.define_name("Sales", "=Sheet1!$G$1:$H$10");

  // Define a local/worksheet name. This overrides the global "Sales" name
  // with a local defined name.
  workbook.define_name("Sheet2!Sales", "=Sheet2!$G$1:$G$10");

  // Write some text to the worksheets and one of the defined name in a formula.
  worksheet1.set_column(0, 0, 45);
  worksheet1.write_string(0, 0, "This worksheet contains some defined names.");
  worksheet1.write_string(1, 0, "See Formulas -> Name Manager above.");
  worksheet1.write_string(2, 0, "Example formula in cell B3 ->");
  worksheet1.write_formula(2, 1, "=Exchange_rate");

  worksheet2.set_column(0, 0, 45);
  worksheet2.write_string(0, 0, "This worksheet contains some defined names.");
  worksheet2.write_string(1, 0, "See Formulas -> Name Manager above.");
  worksheet2.write_string(2, 0, "Example formula in cell B3 ->");
  worksheet2.write_formula(2, 1, "=Exchange_rate");

  workbook.save("defined_name.xlsx");
}
