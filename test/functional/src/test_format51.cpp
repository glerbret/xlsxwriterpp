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

  const double value{123.456};

  worksheet.set_column(0, 0, 12);

  const xwpp::format_t* format1 = workbook.format_builder().num_format("0.0").build();
  worksheet.write(0, 0, value, format1);

  const xwpp::format_t* format2 = workbook.format_builder().num_format("0.000").build();
  worksheet.write(1, 0, value, format2);

  const xwpp::format_t* format3 = workbook.format_builder().num_format("0.0000").build();
  worksheet.write(2, 0, value, format3);

  const xwpp::format_t* format4 = workbook.format_builder().num_format("0.00000").build();
  worksheet.write(3, 0, value, format4);

  workbook.save("test_format51.xlsx");
}
