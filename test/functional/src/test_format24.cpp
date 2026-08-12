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

  format->set_rotation(270);
  format->set_indent(1);
  format->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  format->set_align(xwpp::format_alignments_t::VERTICAL_TOP);

  worksheet.set_row(0, 75);

  worksheet.write_string(0, 0, "ABCD", format);

  workbook.save("test_format24.xlsx");
}
