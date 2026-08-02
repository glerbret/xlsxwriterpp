/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  workbook.unset_default_url_format();

  worksheet.write_url(
    CELL("A1"),
    "http://www.example.com/"
    "some_long_url_that_is_255_characters_long_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_"
    "abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_"
    "abcdefgh_abcdefgh_z#some_long_location_that_is_255_characters_long_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_"
    "abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_"
    "abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_abcdefgh_z");

  workbook.save("test_hyperlink24.xlsx");
}
