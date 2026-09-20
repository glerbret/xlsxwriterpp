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

  worksheet.print_area("A1:F9");

  worksheet.write("A1", "Foo");

  workbook.save("test_print_area06.xlsx");
}
