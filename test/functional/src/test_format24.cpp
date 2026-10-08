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

  const xwpp::format_t* format = workbook.format_builder()
                                   .vertical_text()
                                   .indent(1)
                                   .align(xwpp::format_horizontal_alignments_t::CENTER)
                                   .align(xwpp::format_vertical_alignments_t::TOP)
                                   .build();

  worksheet.set_row(0, 75);

  worksheet.write(0, 0, "ABCD", format);

  workbook.save("test_format24.xlsx");
}
