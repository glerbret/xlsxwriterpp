/*
 * A simple example of some of the features of the Xlsxwriter++ library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format.
  xwpp::format_t* format = workbook.add_format();

  // Set the bold property for the format
  format->set_bold();

  // Change the column width for clarity.
  worksheet.set_column(0, 0, 20);

  // Write some simple text.
  worksheet.write_string(0, 0, "Hello");

  // Text with formatting.
  worksheet.write_string(1, 0, "World", format);

  // Write some numbers.
  worksheet.write_number(2, 0, 123);
  worksheet.write_number(3, 0, 123.456);

  // Insert an image.
  worksheet.insert_image(1, 2, "logo.png");

  workbook.save("demo.xlsx");
}
