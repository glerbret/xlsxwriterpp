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

  const xwpp::format_t* format = workbook.format_builder().align(xwpp::format_horizontal_alignments_t::CENTER).build();

  worksheet.merge_range(1, 1, 1, 2, "Foo", format);
  worksheet.merge_range(1, 3, 1, 4, "Foo", format);
  worksheet.merge_range(1, 5, 1, 6, "Foo", format);

  workbook.save("test_merge_range03.xlsx");
}
