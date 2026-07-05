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

  format1->set_num_format_index(2);
  format2->set_num_format_index(12);

  worksheet.write_number(0, 0, 1.2222);
  worksheet.write_number(1, 0, 1.2222, format1);
  worksheet.write_number(2, 0, 1.2222, format2);
  worksheet.write_number(3, 0, 1.2222);
  worksheet.write_number(4, 0, 1.2222);

  workbook.save("test_data08.xlsx");
}
