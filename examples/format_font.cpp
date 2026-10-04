/*
 * Example of writing some data with font formatting to a simple Excel
 * file using Xlsxwriter++.
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

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 20);

  // Write some formatted strings.
  {
    // Set the bold property for format 1.
    const xwpp::format_t* format = workbook.format_builder().bold().build();

    worksheet.write(0, 0, "This is bold", format);
  }

  {
    // Set the italic property for format 2.
    const xwpp::format_t* format = workbook.format_builder().italic().build();

    worksheet.write(1, 0, "This is italic", format);
  }

  {
    // Set the bold and italic properties for format 3.
    const xwpp::format_t* format = workbook.format_builder().bold().italic().build();

    worksheet.write(2, 0, "Bold and italic", format);
  }

  workbook.save("format_font.xlsx");
}
