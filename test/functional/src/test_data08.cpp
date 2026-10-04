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

  const xwpp::format_t* format1 = workbook.format_builder().num_format_index(2).build();
  const xwpp::format_t* format2 = workbook.format_builder().num_format_index(12).build();

  worksheet.write(0, 0, 1.2222);
  worksheet.write(1, 0, 1.2222, format1);
  worksheet.write(2, 0, 1.2222, format2);
  worksheet.write(3, 0, 1.2222);
  worksheet.write(4, 0, 1.2222);

  workbook.save("test_data08.xlsx");
}
