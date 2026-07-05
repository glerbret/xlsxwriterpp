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

  worksheet.set_column(COLS("C:D"), 10.288);
  worksheet.set_column(COLS("F:G"), 10.288);

  // Add some valid tables.
  worksheet.add_table(RANGE("C2:D3"));

  xwpp::table_options_t options1 = {.name_ = "Table2", .no_header_row_ = true};
  worksheet.add_table(RANGE("F3:G3"), options1);

  // Test incorrect tables. These should be ignored with a warning.
  // TODO Throw an exception, test to rework
  /*
  // Add a table with an incorrect range.
  worksheet.add_table(RANGE("I2:J2"));

  // Check incorrect table names.
  xwpp::table_options_t options2 = {.name_ = "Has space"};
  worksheet.add_table(RANGE("F3:G33"), options2);

  xwpp::table_options_t options3 = {.name_ = "Table["};
  worksheet.add_table(RANGE("F3:G33"), options3);

  xwpp::table_options_t options4 = {
      .name_ =
          "This_is_a_long_table_name_that_exceeds_a_limit_of_255_characters_"
          "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
          "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"};
  worksheet.add_table(RANGE("F3:G33"), options4);

  xwpp::table_options_t options5 = {.name_ = "c"};
  worksheet.add_table(RANGE("F3:G33"), options5);

  xwpp::table_options_t options6 = {.name_ = "R"};
  worksheet.add_table(RANGE("F3:G33"), options6);

  xwpp::table_options_t options7 = {.name_ = ".Table"};
  worksheet.add_table(RANGE("F3:G33"), options7);

  xwpp::table_options_t options8 = {.name_ = "1Table"};
  worksheet.add_table(RANGE("F3:G33"), options8);
*/
  workbook.save("test_table26.xlsx");
}
