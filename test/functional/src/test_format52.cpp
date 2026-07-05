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
  xwpp::format_t* format;

  double value = 123.456;

  worksheet.set_column(0, 0, 12);

  format = workbook.add_format();
  format->set_num_format("0.0");
  worksheet.write_number(0, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.000");
  worksheet.write_number(1, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.0000");
  worksheet.write_number(2, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.00000");
  worksheet.write_number(3, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.0");
  format->set_bold();
  worksheet.write_number(4, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.000");
  format->set_bold();
  worksheet.write_number(5, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.0000");
  format->set_bold();
  worksheet.write_number(6, 0, value, format);

  format = workbook.add_format();
  format->set_num_format("0.00000");
  format->set_bold();
  worksheet.write_number(7, 0, value, format);

  workbook.save("test_format52.xlsx");
}
