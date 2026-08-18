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

  const double value = 123.456;

  worksheet.set_column(0, 0, 12);

  xwpp::format_t* format1 = workbook.add_format();
  format1->set_num_format("0.0");
  worksheet.write_number(0, 0, value, format1);

  xwpp::format_t* format2 = workbook.add_format();
  format2->set_num_format("0.000");
  worksheet.write_number(1, 0, value, format2);

  xwpp::format_t* format3 = workbook.add_format();
  format3->set_num_format("0.0000");
  worksheet.write_number(2, 0, value, format3);

  xwpp::format_t* format4 = workbook.add_format();
  format4->set_num_format("0.00000");
  worksheet.write_number(3, 0, value, format4);

  xwpp::format_t* format5 = workbook.add_format();
  format5->set_num_format("0.0");
  format5->set_bold();
  worksheet.write_number(4, 0, value, format5);

  xwpp::format_t* format6 = workbook.add_format();
  format6->set_num_format("0.000");
  format6->set_bold();
  worksheet.write_number(5, 0, value, format6);

  xwpp::format_t* format7 = workbook.add_format();
  format7->set_num_format("0.0000");
  format7->set_bold();
  worksheet.write_number(6, 0, value, format7);

  xwpp::format_t* format8 = workbook.add_format();
  format8->set_num_format("0.00000");
  format8->set_bold();
  worksheet.write_number(7, 0, value, format8);

  workbook.save("test_format52.xlsx");
}
