/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_paper(xwpp::paper_size_t::A4);
  worksheet.set_dpi(0, 200);

  worksheet.gridlines(xwpp::gridlines_t::SHOW_PRINT_GRIDLINES);
  worksheet.print_row_col_headers();
  worksheet.center_horizontally();
  worksheet.center_vertically();

  worksheet.write("A1", "Foo");

  workbook.save("test_print_options05.xlsx");
}
