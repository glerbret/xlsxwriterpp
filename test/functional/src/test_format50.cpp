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

  format1->set_num_format("#,##0.00000");
  format2->set_num_format("#,##0.0");

  worksheet.set_column(0, 0, 12);

  worksheet.write_number(0, 0, 1234.5, format1);
  worksheet.write_number(1, 0, 1234.5, format2);

  workbook.save("test_format50.xlsx");
}
