/*
 * Examples of how to add data validation and dropdown lists using the
 * Xlsxwriter++ library.
 *
 * Data validation is a feature of Excel which allows you to restrict the data
 * that a user enters in a cell and to display help and warning messages. It
 * also allows you to restrict input to values in a dropdown list.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <chrono>
#include <string>
#include <vector>

using namespace std::literals::chrono_literals;

namespace
{

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* format)
{
  worksheet.write_string(CELL("A1"), "Some examples of data validation in Xlsxwriter++", format);
  worksheet.write_string(CELL("B1"), "Enter values in this column", format);
  worksheet.write_string(CELL("D1"), "Sample Data", format);

  worksheet.write_string(CELL("D3"), "Integers");
  worksheet.write_number(CELL("E3"), 1);
  worksheet.write_number(CELL("F3"), 10);

  worksheet.write_string(CELL("D4"), "List data");
  worksheet.write_string(CELL("E4"), "open");
  worksheet.write_string(CELL("F4"), "high");
  worksheet.write_string(CELL("G4"), "close");

  worksheet.write_string(CELL("D5"), "Formula");
  worksheet.write_formula(CELL("E5"), "=AND(F5=50,G5=60)");
  worksheet.write_number(CELL("F5"), 50);
  worksheet.write_number(CELL("G5"), 60);
}

}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  xwpp::data_validation_t data_validation;

  // Add a format to use to highlight the header cells.
  xwpp::format_t* format = workbook.add_format();
  format->set_border(xwpp::format_borders_t::THIN);
  format->set_fg_color(static_cast<xwpp::color_t>(0xC6EFCE));
  format->set_bold();
  format->set_text_wrap();
  format->set_align(xwpp::format_alignments_t::VERTICAL_CENTER);
  format->set_indent(1);

  // Write some data for the validations.
  write_worksheet_data(worksheet, format);

  // Set up layout of the worksheet.
  worksheet.set_column(0, 0, 55);
  worksheet.set_column(1, 1, 15);
  worksheet.set_column(3, 3, 15);
  worksheet.set_row(0, 36);

  // Example 1. Limiting input to an integer in a fixed range.
  worksheet.write_string(CELL("A3"), "Enter an integer between 1 and 10");

  data_validation.validate_       = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_       = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_number_ = 1;
  data_validation.maximum_number_ = 10;

  worksheet.data_validation_cell(CELL("B3"), data_validation);

  // Example 2. Limiting input to an integer outside a fixed range.
  worksheet.write_string(CELL("A5"), "Enter an integer that is not between 1 and 10 (using cell references)");

  data_validation.validate_        = xwpp::validation_types_t::INTEGER_FORMULA;
  data_validation.criteria_        = xwpp::validation_criteria_t::NOT_BETWEEN;
  data_validation.minimum_formula_ = "=E3";
  data_validation.maximum_formula_ = "=F3";

  worksheet.data_validation_cell(CELL("B5"), data_validation);

  // Example 3. Limiting input to an integer greater than a fixed value.
  worksheet.write_string(CELL("A7"), "Enter an integer greater than 0");

  data_validation.validate_     = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_     = xwpp::validation_criteria_t::GREATER_THAN;
  data_validation.value_number_ = 0;

  worksheet.data_validation_cell(CELL("B7"), data_validation);

  // Example 4. Limiting input to an integer less than a fixed value.
  worksheet.write_string(CELL("A9"), "Enter an integer less than 10");

  data_validation.validate_     = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_     = xwpp::validation_criteria_t::LESS_THAN;
  data_validation.value_number_ = 10;

  worksheet.data_validation_cell(CELL("B9"), data_validation);

  // Example 5. Limiting input to a decimal in a fixed range.
  worksheet.write_string(CELL("A11"), "Enter a decimal between 0.1 and 0.5");

  data_validation.validate_       = xwpp::validation_types_t::DECIMAL;
  data_validation.criteria_       = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_number_ = 0.1;
  data_validation.maximum_number_ = 0.5;

  worksheet.data_validation_cell(CELL("B11"), data_validation);

  // Example 6. Limiting input to a value in a dropdown list.
  worksheet.write_string(CELL("A13"), "Select a value from a dropdown list");

  const std::vector<std::string> list{"open", "high", "close"};
  data_validation.validate_   = xwpp::validation_types_t::LIST;
  data_validation.value_list_ = list;

  worksheet.data_validation_cell(CELL("B13"), data_validation);

  // Example 7. Limiting input to a value in a dropdown list.
  worksheet.write_string(CELL("A15"), "Select a value from a dropdown list (using a cell range)");

  data_validation.validate_      = xwpp::validation_types_t::LIST_FORMULA;
  data_validation.value_formula_ = "=$E$4:$G$4";

  worksheet.data_validation_cell(CELL("B15"), data_validation);

  // Example 8. Limiting input to a date in a fixed range.
  worksheet.write_string(CELL("A17"), "Enter a date between 1/1/2024 and 12/12/2024");

  data_validation.validate_         = xwpp::validation_types_t::DATE;
  data_validation.criteria_         = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_datetime_ = std::chrono::sys_days{2024y / std::chrono::January / 1d} + 0h + 0min + 0s;
  data_validation.maximum_datetime_ = std::chrono::sys_days{2024y / std::chrono::December / 12d} + 0h + 0min + 0s;

  worksheet.data_validation_cell(CELL("B17"), data_validation);

  // Example 9. Limiting input to a time in a fixed range.
  worksheet.write_string(CELL("A19"), "Enter a time between 6:00 and 12:00");

  data_validation.validate_         = xwpp::validation_types_t::TIME;
  data_validation.criteria_         = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_datetime_ = std::chrono::system_clock::time_point{6h + 0min + 0s};
  data_validation.maximum_datetime_ = std::chrono::system_clock::time_point{12h + 0min + 0s};

  worksheet.data_validation_cell(CELL("B19"), data_validation);

  // Example 10. Limiting input to a string greater than a fixed length.
  worksheet.write_string(CELL("A21"), "Enter a string longer than 3 characters");

  data_validation.validate_     = xwpp::validation_types_t::LENGTH;
  data_validation.criteria_     = xwpp::validation_criteria_t::GREATER_THAN;
  data_validation.value_number_ = 3;

  worksheet.data_validation_cell(CELL("B21"), data_validation);

  // Example 11. Limiting input based on a formula.
  worksheet.write_string(CELL("A23"), "Enter a value if the following is true \"=AND(F5=50,G5=60)\"");

  data_validation.validate_      = xwpp::validation_types_t::CUSTOM_FORMULA;
  data_validation.value_formula_ = "=AND(F5=50,G5=60)";

  worksheet.data_validation_cell(CELL("B23"), data_validation);

  // Example 12. Displaying and modifying data validation messages.
  worksheet.write_string(CELL("A25"), "Displays a message when you select the cell");

  data_validation.validate_       = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_       = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_number_ = 1;
  data_validation.maximum_number_ = 100;
  data_validation.input_title_    = "Enter an integer:";
  data_validation.input_message_  = "between 1 and 100";

  worksheet.data_validation_cell(CELL("B25"), data_validation);

  // Example 13. Displaying and modifying data validation messages.
  worksheet.write_string(CELL("A27"), "Display a custom error message when integer isn't between 1 and 100");

  data_validation.validate_       = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_       = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_number_ = 1;
  data_validation.maximum_number_ = 100;
  data_validation.input_title_    = "Enter an integer:";
  data_validation.input_message_  = "between 1 and 100";
  data_validation.error_title_    = "Input value is not valid!";
  data_validation.error_message_  = "It should be an integer between 1 and 100";

  worksheet.data_validation_cell(CELL("B27"), data_validation);

  // Example 14. Displaying and modifying data validation messages.
  worksheet.write_string(CELL("A29"), "Display a custom info message when integer isn't between 1 and 100");

  data_validation.validate_       = xwpp::validation_types_t::INTEGER;
  data_validation.criteria_       = xwpp::validation_criteria_t::BETWEEN;
  data_validation.minimum_number_ = 1;
  data_validation.maximum_number_ = 100;
  data_validation.input_title_    = "Enter an integer:";
  data_validation.input_message_  = "between 1 and 100";
  data_validation.error_title_    = "Input value is not valid!";
  data_validation.error_message_  = "It should be an integer between 1 and 100";
  data_validation.error_type_     = xwpp::validation_error_types_t::INFORMATION;

  worksheet.data_validation_cell(CELL("B29"), data_validation);

  workbook.save("data_validate.xlsx");
}
