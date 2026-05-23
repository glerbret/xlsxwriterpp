/*
 * Example of writing some data to a simple Excel file using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t wb;
  xwpp::worksheet_t& ws = wb.add_worksheet();

  ws.write_string(0, 0, "Hello");
  ws.write_number(1, 0, 123);

  wb.save("hello_world.xlsx");

  return 0;
}
