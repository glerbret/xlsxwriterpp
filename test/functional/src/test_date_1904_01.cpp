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

  xwpp::format_t* format = workbook.add_format();
  format->set_num_format_index(14);

  const xwpp::datetime_t datetime1{1900, 1, 1, 0, 0, 0.0};
  const xwpp::datetime_t datetime2{1902, 9, 26, 0, 0, 0.0};
  const xwpp::datetime_t datetime3{1913, 9, 8, 0, 0, 0.0};
  const xwpp::datetime_t datetime4{1927, 5, 18, 0, 0, 0.0};
  const xwpp::datetime_t datetime5{2173, 10, 14, 0, 0, 0.0};
  const xwpp::datetime_t datetime6{4637, 11, 26, 0, 0, 0.0};

  worksheet.set_column(0, 0, 12);

  worksheet.write_datetime(CELL("A1"), datetime1, format);
  worksheet.write_datetime(CELL("A2"), datetime2, format);
  worksheet.write_datetime(CELL("A3"), datetime3, format);
  worksheet.write_datetime(CELL("A4"), datetime4, format);
  worksheet.write_datetime(CELL("A5"), datetime5, format);
  worksheet.write_datetime(CELL("A6"), datetime6, format);

  workbook.save("test_date_1904_01.xlsx");
}
