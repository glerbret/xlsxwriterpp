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
  xwpp::format_t* num_format   = workbook.add_format();

  num_format->set_num_format(R"([Red]0.0%\ "a")");

  worksheet.set_column(0, 0, 14);

  worksheet.write_number(CELL("A1"), 123, num_format);

  workbook.save("test_escapes06.xlsx");
}
