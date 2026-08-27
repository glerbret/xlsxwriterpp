/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  // TODO Currently use_1904_epoch must be called before creating worksheet.
  // Must be provides as workbook option on creation or push to worksheet on file encoding
  workbook.use_1904_epoch();
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format = workbook.add_format();
  format->set_num_format_index(14);

  const xwpp::datetime_t datetime1{.year_ = 1904, .month_ = 1, .day_ = 1, .hour_ = 0, .min_ = 0, .sec_ = 0.0};
  const xwpp::datetime_t datetime2{.year_ = 1906, .month_ = 9, .day_ = 27, .hour_ = 0, .min_ = 0, .sec_ = 0.0};
  const xwpp::datetime_t datetime3{.year_ = 1917, .month_ = 9, .day_ = 9, .hour_ = 0, .min_ = 0, .sec_ = 0.0};
  const xwpp::datetime_t datetime4{.year_ = 1931, .month_ = 5, .day_ = 19, .hour_ = 0, .min_ = 0, .sec_ = 0.0};
  const xwpp::datetime_t datetime5{.year_ = 2177, .month_ = 10, .day_ = 15, .hour_ = 0, .min_ = 0, .sec_ = 0.0};
  const xwpp::datetime_t datetime6{.year_ = 4641, .month_ = 11, .day_ = 27, .hour_ = 0, .min_ = 0, .sec_ = 0.0};

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime(CELL("A1"), datetime1, format);
  worksheet.write_datetime(CELL("A2"), datetime2, format);
  worksheet.write_datetime(CELL("A3"), datetime3, format);
  worksheet.write_datetime(CELL("A4"), datetime4, format);
  worksheet.write_datetime(CELL("A5"), datetime5, format);
  worksheet.write_datetime(CELL("A6"), datetime6, format);

  workbook.save("test_date_1904_02.xlsx");
}