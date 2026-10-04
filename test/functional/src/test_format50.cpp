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

  const xwpp::format_t* format1 = workbook.format_builder().num_format("#,##0.00000").build();
  const xwpp::format_t* format2 = workbook.format_builder().num_format("#,##0.0").build();

  worksheet.set_column(0, 0, 12);

  worksheet.write(0, 0, 1234.5, format1);
  worksheet.write(1, 0, 1234.5, format2);

  workbook.save("test_format50.xlsx");
}
