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

  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();

  worksheet.set_row(0, 30);

  format1->set_font_name("Arial");
  format1->set_bold();
  format1->set_align(xwpp::format_alignments_t::HORIZONTAL_LEFT);
  format1->set_align(xwpp::format_alignments_t::VERTICAL_BOTTOM);

  format2->set_font_name("Arial");
  format2->set_bold();
  format2->set_rotation(90);
  format2->set_align(xwpp::format_alignments_t::HORIZONTAL_CENTER);
  format2->set_align(xwpp::format_alignments_t::VERTICAL_BOTTOM);

  worksheet.write_string(0, 0, "Foo", format1);
  worksheet.write_string(0, 1, "Bar", format2);

  workbook.save("test_format02.xlsx");
}
