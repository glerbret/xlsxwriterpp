/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  try
  {
    xwpp::workbook_t workbook;
    workbook.set_max_url_length(255);
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    worksheet.write_string(CELL("A1"), "Foo");
    worksheet.write_string(CELL("A3"), "Bar");

    // TODO Throw an exception
    // This link is too long and should be ignored, with a warning.
    worksheet.write_url(
      CELL("A2"), "http://foo.com/"
                  "this_is_a_long_hyperlink_that_exceeds_a_limit_of_255_characters_"
                  "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                  "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");

    workbook.save("test_hyperlink46.xlsx");
  } catch(const xwpp::xwpp_exception_t&)
  {
    return EXIT_FAILURE;
  }
}
