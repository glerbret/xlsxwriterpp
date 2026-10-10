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

  const xwpp::format_t* pattern = workbook.format_builder().pattern(xwpp::format_patterns_t::MEDIUM_GRAY).build();

  worksheet.write("A1", "", pattern);

  workbook.save("test_format16.xlsx");
}
