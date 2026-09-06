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
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  const std::vector<std::string> list1{"Foo", "Bar", "Baz"};
  const std::vector<std::string> list2{"Foobar", "Foobas", "Foobat", "Foobau", "Foobav", "Foobaw", "Foobax", "Foobay",
                                       "Foobaz", "Foobba", "Foobbb", "Foobbc", "Foobbd", "Foobbe", "Foobbf", "Foobbg",
                                       "Foobbh", "Foobbi", "Foobbj", "Foobbk", "Foobbl", "Foobbm", "Foobbn", "Foobbo",
                                       "Foobbp", "Foobbq", "Foobbr", "Foobbs", "Foobbt", "Foobbu", "Foobbv", "Foobbw",
                                       "Foobbx", "Foobby", "Foobbz", "Foobca", "End"};

  const xwpp::data_validation_t data_validation1{
    .validate_      = xwpp::validation_types_t::LIST,
    .value_list_    = list1,
    .input_title_   = "This is the input title",
    .input_message_ = "This is the input message",
  };
  worksheet.data_validation_cell(CELL("C2"), data_validation1);

  const xwpp::data_validation_t data_validation2{
    .validate_    = xwpp::validation_types_t::LIST,
    .value_list_  = list2,
    .input_title_ = "This is the longest input title1",
    .input_message_ =
      "This is the longest input message "
      "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
  };
  worksheet.data_validation_cell(CELL("D6"), data_validation2);

  workbook.save("test_data_validation05.xlsx");
}
