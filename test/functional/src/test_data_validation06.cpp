/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet         = workbook.add_worksheet();
  const std::vector<std::string> list1 = {"Foo", "Bar", "Baz"};
  const std::vector<std::string> list2 = {
    "Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax", "Foobay", "Foobaz", "Foobba",
    "Foobbb", "Foobbc", "Foobbd", "Foobbe", "Foobbf", "Foobbg", "Foobbh", "Foobbi", "Foobbj", "Foobbk",
    "Foobbl", "Foobbm", "Foobbn", "Foobbo", "Foobbp", "Foobbq", "Foobbr", "Foobbs", "Foobbt", "Foobbu",
    "Foobbv", "Foobbw", "Foobbx", "Foobby", "Foobbz", "Foobca", "End1"};

  xwpp::data_validation_t data_validation{
    .validate_      = xwpp::validation_types_t::LIST,
    .value_list_    = list1,
    .input_title_   = "This is the input title",
    .input_message_ = "This is the input message",
  };
  worksheet.data_validation_cell(CELL("C2"), data_validation);

  data_validation.validate_    = xwpp::validation_types_t::LIST;
  // Next string exceeds Excel's limit so validation should be ignored.
  data_validation.value_list_  = list2;
  data_validation.input_title_ = "This is the longest input title1";
  data_validation.input_message_ =
    "This is the longest input message "
    "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
    "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
  worksheet.data_validation_cell(CELL("D6"), data_validation);

  workbook.save("test_data_validation06.xlsx");
}
