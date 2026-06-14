/*
 * An example of embedding images into a worksheet using the Xlsxwriter++
 * library.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

int main()
{
  // Create a new workbook and add a worksheet.
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Change some of the column widths for clarity.
  worksheet.set_column(COLS("A:B"), 30);

  // Embed an image.
  worksheet.write_string(CELL("A2"), "Embed an image in a cell:");
  worksheet.embed_image(CELL("B2"), "logo.png");

  // Make a row bigger and embed the image.
  worksheet.set_row(3, 72);
  worksheet.write_string(CELL("A4"), "Embed an image in a cell:");
  worksheet.embed_image(CELL("B4"), "logo.png");

  // Make a row bigger and embed the image.
  worksheet.set_row(5, 150);
  worksheet.write_string(CELL("A6"), "Embed an image in a cell:");
  worksheet.embed_image(CELL("B6"), "logo.png");

  workbook.save("embed_images.xlsx");
}
