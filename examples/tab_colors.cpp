/*
 * Example of how to set Excel worksheet tab colors using Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet();

  // Set the tab colors.
  worksheet1.set_tab_color(xwpp::color_t::RED);
  worksheet2.set_tab_color(xwpp::color_t::GREEN);
  worksheet3.set_tab_color(static_cast<xwpp::color_t>(0xFF9900)); /* Orange. */

  // worksheet4 will have the default color.
  worksheet4.write_string(0, 0, "Hello");

  workbook.save("tab_colors.xlsx");
}
