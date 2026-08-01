/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  std::vector<xwpp::row_num_t> breaks;
  breaks.reserve(1028);
  for(xwpp::row_num_t i = 0; i < 1027; i++)
  {
    breaks.push_back(i + 1);
  }

  worksheet.set_paper(9);
  worksheet.set_dpi(0, 200);

  worksheet.set_h_pagebreaks(breaks);

  worksheet.write_string(CELL("A1"), "Foo", nullptr);

  workbook.save("test_page_breaks03.xlsx");
}
