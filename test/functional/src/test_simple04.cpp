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

  const xwpp::format_t* format1 = workbook.format_builder().num_format_index(20).build();
  const xwpp::format_t* format2 = workbook.format_builder().num_format_index(14).build();

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime(0, 0, xwpp::datetime_t{.year = 0, .month = 0, .day = 0, .hour = 12, .min = 0, .sec = 0.},
                           format1);
  worksheet.write_datetime(1, 0, xwpp::datetime_t{.year = 2013, .month = 1, .day = 27, .hour = 0, .min = 0, .sec = 0.},
                           format2);

  workbook.save("test_simple04.xlsx");
}
