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

  const xwpp::format_t* format1 = workbook.format_builder()
                                    .font_name("Arial")
                                    .bold()
                                    .align(xwpp::format_alignments_t::HORIZONTAL_LEFT)
                                    .align(xwpp::format_alignments_t::VERTICAL_BOTTOM)
                                    .build();
  const xwpp::format_t* format2 = workbook.format_builder()
                                    .font_name("Arial")
                                    .bold()
                                    .rotation(90)
                                    .align(xwpp::format_alignments_t::HORIZONTAL_CENTER)
                                    .align(xwpp::format_alignments_t::VERTICAL_BOTTOM)
                                    .build();

  worksheet.set_row(0, 30);

  worksheet.write(0, 0, "Foo", format1);
  worksheet.write(0, 1, "Bar", format2);

  workbook.save("test_format02.xlsx");
}
