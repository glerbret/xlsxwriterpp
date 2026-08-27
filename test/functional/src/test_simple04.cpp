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

  worksheet.write_datetime(
    0, 0, xwpp::datetime_t{.year_ = 0, .month_ = 0, .day_ = 0, .hour_ = 12, .min_ = 0, .sec_ = 0.}, format1);
  worksheet.write_datetime(
    1, 0, xwpp::datetime_t{.year_ = 2013, .month_ = 1, .day_ = 27, .hour_ = 0, .min_ = 0, .sec_ = 0.}, format2);

  workbook.save("test_simple04.xlsx");
}
