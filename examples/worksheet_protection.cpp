/*
 * Example of cell locking and formula hiding in an Excel worksheet using
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

  xwpp::format_t* unlocked = workbook.add_format();
  unlocked->set_unlocked();

  xwpp::format_t* hidden = workbook.add_format();
  hidden->set_hidden();

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 40);

  // Turn worksheet protection on without a password.
  worksheet.protect();

  // Write a locked, unlocked and hidden cell.
  worksheet.write_string(0, 0, "B1 is locked. It cannot be edited.");
  worksheet.write_string(1, 0, "B2 is unlocked. It can be edited.");
  worksheet.write_string(2, 0, "B3 is hidden. The formula isn't visible.");

  worksheet.write_formula(0, 1, "=1+2"); /* Locked by default. */
  worksheet.write_formula(1, 1, "=1+2", unlocked);
  worksheet.write_formula(2, 1, "=1+2", hidden);

  workbook.save("protection.xlsx");
}
