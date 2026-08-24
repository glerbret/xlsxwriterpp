/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>

using namespace std::literals::chrono_literals;

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();
  format1->set_num_format_index(20);
  format2->set_num_format_index(14);

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime(0, 0, std::chrono::system_clock::time_point{12h + 0min + 0s}, format1);
  worksheet.write_datetime(1, 0, std::chrono::sys_days{2013y / std::chrono::January / 27d} + 0h + 0min + 0s, format2);

  workbook.save("test_simple40.xlsx");
}
