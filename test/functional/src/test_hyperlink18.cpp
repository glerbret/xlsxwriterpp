/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  // Set shorter length for testing.
  workbook.set_max_url_length(255);
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  workbook.unset_default_url_format();

  worksheet.write_url(CELL("A1"),
                      "http://google.com/"
                      "000000000011111111112222222222333333333344444444445555555555666666666667777777777788888888889999"
                      "999999900000000001111111111222222222233333333334444444444555555555566666666666777777777778888888"
                      "88899999999999000000000011111111112222222222x");

  // TODO Throw an exception (contrary to C libxlsxwriter that ignored it)
  // TODO Dedicated test to add
  // This longer url should be ignored.
  /*
    worksheet.write_url(CELL("A1"),
                        "http://google.com/"
                        "000000000011111111112222222222333333333344444444445555555555666666666667777777777788888888889999"
                        "999999900000000001111111111222222222233333333334444444444555555555566666666666777777777778888888"
                        "88899999999999000000000011111111112222222222xy");
  */

  workbook.save("test_hyperlink18.xlsx");
}
