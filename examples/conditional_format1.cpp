/*
 * An a simple example of how to add conditional formatting to an
 * Xlsxwriter++ file.
 *
 * See conditional_format.c for a more comprehensive example.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some sample data.
  worksheet.write("B1", 34);
  worksheet.write("B2", 32);
  worksheet.write("B3", 31);
  worksheet.write("B4", 35);
  worksheet.write("B5", 36);
  worksheet.write("B6", 30);
  worksheet.write("B7", 38);
  worksheet.write("B8", 38);
  worksheet.write("B9", 32);

  // Add a format with red text.
  const xwpp::format_t* custom_format = workbook.format_builder().font_color(xwpp::color_t::red()).build();

  // Create a conditional format object. A static object would also work.
  const xwpp::conditional_format_t conditional_format{
    .type_     = xwpp::conditional_format_types_t::CELL,
    .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
    .value_    = 33,
    .format_   = custom_format,
  };

  // Now apply the format to data range.
  worksheet.conditional_format_range("B1:B9", conditional_format);

  workbook.save("conditional_format_simple.xlsx");
}
