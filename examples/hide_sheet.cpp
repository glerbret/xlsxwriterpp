/*
 * Example of how to hide a worksheet using libxlsxwriter.
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
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  // Hide Sheet2. It won't be visible until it is unhidden in Excel.
  worksheet2.hide();

  worksheet1.write_string(0, 0, "Sheet2 is hidden");
  worksheet2.write_string(0, 0, "Now it's my turn to find you!");
  worksheet3.write_string(0, 0, "Sheet2 is hidden");

  // Make the first column wider to make the text clearer.
  worksheet1.set_column(0, 0, 30);
  worksheet2.set_column(0, 0, 30);
  worksheet3.set_column(0, 0, 30);

  workbook.save("hide_sheet.xlsx");
}
