/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();
  xwpp::format_t* format3 = workbook.add_format();

  format1->set_bold();

  format2->set_italic();

  format3->set_bold();
  format3->set_italic();

  worksheet.write_string(CELL("A1"), "Foo", format1);
  worksheet.write_string(CELL("A2"), "Bar", format2);
  worksheet.write_string(CELL("A3"), "Baz", format3);

  workbook.save("test_data06.xlsx");
}
