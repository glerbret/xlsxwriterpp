/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // TODO Error conditions. Throw exception not ignored as C libxlsxwriter
  // Disable for the moment, dedicated tests must be added
  /*
    // Test empty string.
    worksheet.set_header("");
    worksheet.set_footer("");

    // Test long string.
    const str::sting long_string = "This is a long string longer than the Excel limit of 255 characters used for testing
    " "                                                                                                " " end.";
    worksheet.set_header(long_string);
    worksheet.set_footer(long_string);

    // Test &G placeholder without image.
    worksheet.set_header("&L&G");
    worksheet.set_footer("&L&G");
    worksheet.set_header("&L&G", {});
    worksheet.set_footer("&L&G", {});

    // Test too many images/not enough placeholders.
    xwpp::header_footer_options_t options3 = {.image_left_ = "images/red.jpg", .image_right_ = "images/red.jpg"};
    worksheet.set_header("&L&G", options3);
    worksheet.set_footer("&L&G", options3);
  */
  // Write a non-error header.
  const xwpp::header_footer_options_t header_options{.image_left_ = "images/red.jpg"};

  worksheet.set_header("&L&G", header_options);

  workbook.save("test_header_image01.xlsx");
}
