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

  worksheet.fit_to_pages(1, 1);
  worksheet.set_paper(xwpp::paper_size_t::A4);
  worksheet.set_dpi(0, 200);

  worksheet.write(CELL("A1"), "Foo");

  workbook.save("test_fit_to_pages01.xlsx");
}
