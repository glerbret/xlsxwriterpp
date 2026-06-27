/*
 * Example of writing urls/hyperlinks with the Xlsxwriter++ library.
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

  // Get the default url format (used in the overwriting examples below).
  xwpp::format_t* url_format = workbook.get_default_url_format();

  // Create a user defined link format.
  xwpp::format_t* red_format = workbook.add_format();
  red_format->set_underline(xwpp::format_underlines_t::SINGLE);
  red_format->set_font_color(xwpp::color_t::RED);

  // Widen the first column to make the text clearer.
  worksheet.set_column(0, 0, 30);

  // Write a hyperlink. A default blue underline will be used if the format is not present.
  worksheet.write_url(0, 0, "https://github.com/glerbret/xlsxwriterpp");

  // Write a hyperlink but overwrite the displayed string. Note, we need to
  // specify the format for the string to match the default hyperlink.
  worksheet.write_url(2, 0, "https://github.com/glerbret/xlsxwriterpp");
  worksheet.write_string(2, 0, "Read the documentation.", url_format);

  // Write a hyperlink with a different format.
  worksheet.write_url(4, 0, "https://github.com/glerbret/xlsxwriterpp", red_format);

  // Write a mail hyperlink.
  worksheet.write_url(6, 0, "mailto:john@doe.com");

  // Write a mail hyperlink and overwrite the displayed string. We again
  // specify the format for the string to match the default hyperlink.
  worksheet.write_url(8, 0, "mailto:john@doe.com");
  worksheet.write_string(8, 0, "Drop me a line.", url_format);

  // Write a hyperlink with a text and a tooltip
  worksheet.write_url(10, 0, "https://github.com/glerbret/xlsxwriterpp", url_format, "With text and tooltip",
                      "tooltip");

  workbook.save("hyperlinks.xlsx");
}
