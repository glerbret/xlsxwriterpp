/*
 * An example of turning off worksheet cells errors/warnings using
 * Xlsxwriter++.
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

  // Write strings that looks like numbers. This will cause an Excel warning.
  worksheet.write_string(CELL("C2"), "123");
  worksheet.write_string(CELL("C3"), "123");

  // Write a divide by zero formula. This will also cause an Excel warning.
  worksheet.write_formula(CELL("C5"), "=1/0");
  worksheet.write_formula(CELL("C6"), "=1/0");

  // Turn off some of the warnings:
  worksheet.ignore_errors(xwpp::ignore_errors_t::NUMBER_STORED_AS_TEXT, "C3");
  worksheet.ignore_errors(xwpp::ignore_errors_t::EVAL_ERROR, "C6");

  // Write some descriptions for the cells and make the column wider for clarity.
  worksheet.set_column(1, 1, 16);
  worksheet.write_string(CELL("B2"), "Warning:");
  worksheet.write_string(CELL("B3"), "Warning turned off:");
  worksheet.write_string(CELL("B5"), "Warning:");
  worksheet.write_string(CELL("B6"), "Warning turned off:");

  workbook.save("ignore_errors.xlsx");
}
