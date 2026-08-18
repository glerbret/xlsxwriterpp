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
  xwpp::worksheet_t& worksheet        = workbook.add_worksheet();
  const std::vector<std::string> list = {"Foo", "Bar", "Baz"};

  const xwpp::data_validation_t data_validation{
    .validate_      = xwpp::validation_types_t::LIST,
    .value_list_    = list,
    .input_title_   = "This is the input title",
    .input_message_ = "This is the input message",
  };

  worksheet.data_validation_cell(CELL("C2"), data_validation);

  workbook.save("test_data_validation02.xlsx");
}
