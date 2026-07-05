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

  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();

  format1->set_bold();
  format2->set_bold();
  format2->set_num_format_index(1);

  worksheet.write_number(CELL("A1"), 1, format1);
  worksheet.write_number(CELL("A2"), 2, format2);

  workbook.save("test_format15.xlsx");
}
