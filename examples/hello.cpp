/*
 * Example of writing some data to a simple Excel file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.write_string(0, 0, "Hello");
  worksheet.write_number(1, 0, 123);

  workbook.save("hello_world.xlsx");
}
