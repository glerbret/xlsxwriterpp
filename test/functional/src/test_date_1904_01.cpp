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

  const xwpp::format_t* format = workbook.format_builder().num_format_index(14).build();
  const xwpp::datetime_t datetime1{.year = 1900, .month = 1, .day = 1, .hour = 0, .min = 0, .sec = 0.0};
  const xwpp::datetime_t datetime2{.year = 1902, .month = 9, .day = 26, .hour = 0, .min = 0, .sec = 0.0};
  const xwpp::datetime_t datetime3{.year = 1913, .month = 9, .day = 8, .hour = 0, .min = 0, .sec = 0.0};
  const xwpp::datetime_t datetime4{.year = 1927, .month = 5, .day = 18, .hour = 0, .min = 0, .sec = 0.0};
  const xwpp::datetime_t datetime5{.year = 2173, .month = 10, .day = 14, .hour = 0, .min = 0, .sec = 0.0};
  const xwpp::datetime_t datetime6{.year = 4637, .month = 11, .day = 26, .hour = 0, .min = 0, .sec = 0.0};

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime("A1", datetime1, format);
  worksheet.write_datetime("A2", datetime2, format);
  worksheet.write_datetime("A3", datetime3, format);
  worksheet.write_datetime("A4", datetime4, format);
  worksheet.write_datetime("A5", datetime5, format);
  worksheet.write_datetime("A6", datetime6, format);

  workbook.save("test_date_1904_01.xlsx");
}
