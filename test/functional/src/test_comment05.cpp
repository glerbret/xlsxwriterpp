/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  /* xwpp::worksheet_t& worksheet2 = */ workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();

  for(xwpp::row_num_t row_num{0}; row_num <= 127; row_num++)
  {
    for(xwpp::col_num_t col_num{0}; col_num <= 15; col_num++)
    {
      worksheet1.write_comment(row_num, col_num, "Some text");
    }
  }

  worksheet3.write_comment(CELL("A1"), "More text");

  worksheet1.set_comments_author("John");
  worksheet3.set_comments_author("John");

  workbook.save("test_comment05.xlsx");
}
