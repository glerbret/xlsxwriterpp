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

  xwpp::format_t* format = workbook.add_format();
  format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  format->set_bold();

  worksheet.merge_range(1, 1, 1, 3, "Foo", format);

  workbook.save("test_merge_range04.xlsx");
}
