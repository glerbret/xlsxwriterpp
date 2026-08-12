/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::format_t* format1 = workbook.add_format();
  xwpp::format_t* format2 = workbook.add_format();
  xwpp::format_t* format3 = workbook.add_format();

  format1->set_num_format("0.00;[Red]0.00");
  format2->set_num_format("0.00_ ;\\-0.00\\ ");
  format3->set_num_format("0.00_ ;[Red]\\-0.00\\ ");

  // We manually set the indices to get the same order as the target file.
  format3->get_dxf_index_(format3);
  format2->get_dxf_index_(format2);
  format1->get_dxf_index_(format1);

  worksheet.set_column(COLS("C:F"), 10.288);

  const std::vector<xwpp::table_column_t> columns{{}, {.format_ = format1}, {.format_ = format2}, {.format_ = format3}};
  const xwpp::table_options_t options{.columns_ = columns};

  worksheet.add_table(RANGE("C2:F6"), options);

  worksheet.write_string(2, 2, "Foo");
  worksheet.write_string(3, 2, "Bar");
  worksheet.write_string(4, 2, "Baz");
  worksheet.write_string(5, 2, "Bop");

  worksheet.write_number(2, 3, 1234, format1);
  worksheet.write_number(3, 3, 1256, format1);
  worksheet.write_number(4, 3, 2234, format1);
  worksheet.write_number(5, 3, 1324, format1);

  worksheet.write_number(2, 4, 2000, format2);
  worksheet.write_number(3, 4, 4000, format2);
  worksheet.write_number(4, 4, 3000, format2);
  worksheet.write_number(5, 4, 1000, format2);

  worksheet.write_number(2, 5, 4321, format3);
  worksheet.write_number(3, 5, 4320, format3);
  worksheet.write_number(4, 5, 4332, format3);
  worksheet.write_number(5, 5, 4333, format3);

  workbook.save("test_table14.xlsx");
}
