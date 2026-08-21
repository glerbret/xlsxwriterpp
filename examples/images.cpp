/*
 * An example of inserting images into a worksheet using the Xlsxwriter++
 * library.
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

  // Change some of the column widths for clarity.
  worksheet.set_column(0, 0, 30);

  // Insert an image.
  worksheet.write_string(CELL("A2"), "Insert an image in a cell:");

  worksheet.insert_image(CELL("B2"), "logo.png");

  // Insert an image offset in the cell.
  worksheet.write_string(CELL("A12"), "Insert an offset image:");

  {
    const xwpp::image_options_t options{.x_offset_ = 15, .y_offset_ = 10};
    worksheet.insert_image(CELL("B12"), "logo.png", options);
  }

  // Insert an image with scaling.
  worksheet.write_string(CELL("A22"), "Insert a scaled image:");

  {
    const xwpp::image_options_t options{.x_scale_ = 0.5, .y_scale_ = 0.5};
    worksheet.insert_image(CELL("B22"), "logo.png", options);
  }

  // Insert an image with a hyperlink.
  worksheet.write_string(CELL("A32"), "Insert an image with a hyperlink:");

  {
    const xwpp::image_options_t options{.url_ = "https://github.com/glerbret"};
    worksheet.insert_image(CELL("B32"), "logo.png", options);
  }

  workbook.save("images.xlsx");
}
