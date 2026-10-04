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
  const xwpp::format_t* format2 = workbook.format_builder().italic().build();
  const xwpp::format_t* format3 = workbook.format_builder().bold().italic().build();

  worksheet.write("A1", "Foo", format1);
  worksheet.write("A2", "Bar", format2);
  worksheet.write("A3", "Baz", format3);

  workbook.save("test_data06.xlsx");
}
