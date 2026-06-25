/*
 * An a simple example of how to add conditional formatting to an
 * Xlsxwriter++ file.
 *
 * See conditional_format.c for a more comprehensive example.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Write some sample data.
  worksheet.write_number(CELL("B1"), 34);
  worksheet.write_number(CELL("B2"), 32);
  worksheet.write_number(CELL("B3"), 31);
  worksheet.write_number(CELL("B4"), 35);
  worksheet.write_number(CELL("B5"), 36);
  worksheet.write_number(CELL("B6"), 30);
  worksheet.write_number(CELL("B7"), 38);
  worksheet.write_number(CELL("B8"), 38);
  worksheet.write_number(CELL("B9"), 32);

  // Add a format with red text.
  xwpp::format_t* custom_format = workbook.add_format();
  custom_format->set_font_color(xwpp::color_t::RED);

  // Create a conditional format object. A static object would also work.
  xwpp::conditional_format_t conditional_format{
      .type_     = xwpp::conditional_format_types_t::CELL,
      .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
      .value_    = 33,
      .format_   = custom_format,
  };

  // Now apply the format to data range.
  worksheet.conditional_format_range(RANGE("B1:B9"), conditional_format);

  workbook.save("conditional_format_simple.xlsx");
}
