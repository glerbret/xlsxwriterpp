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

  const xwpp::format_t* format1 = workbook.format_builder().bold().build();
  const xwpp::format_t* format2 = workbook.format_builder().bold().num_format_index(1).build();

  worksheet.write("A1", 1, format1);
  worksheet.write("A2", 2, format2);

  workbook.save("test_format15.xlsx");
}
