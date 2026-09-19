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
  worksheet.write("A1", "Some examples of data validation in Xlsxwriter++", format);
  worksheet.write("B1", "Enter values in this column", format);
  worksheet.write("D1", "Sample Data", format);

  worksheet.write("D3", "Integers");
  worksheet.write("E3", 1);
  worksheet.write("F3", 10);

  worksheet.write("D4", "List data");
  worksheet.write("E4", "open");
  worksheet.write("F4", "high");
  worksheet.write("G4", "close");

  worksheet.write("D5", "Formula");
  worksheet.write_formula("E5", "=AND(F5=50,G5=60)");
  worksheet.write("F5", 50);
  worksheet.write("G5", 60);
}

}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  // Add a format to use to highlight the header cells.
  xwpp::format_t* format = workbook.add_format();
  format->set_border(xwpp::format_borders_t::THIN);
  format->set_fg_color(xwpp::color_t{0xC6EFCE});
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
  {
    worksheet.write("A3", "Enter an integer between 1 and 10");

    const xwpp::data_validation_t data_validation{
      .validate_       = xwpp::validation_types_t::INTEGER,
      .criteria_       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number_ = 1,
      .maximum_number_ = 10,
    };
    worksheet.data_validation_cell("B3", data_validation);
  }

  // Example 2. Limiting input to an integer outside a fixed range.
  {
    worksheet.write("A5", "Enter an integer that is not between 1 and 10 (using cell references)");

    const xwpp::data_validation_t data_validation{
      .validate_        = xwpp::validation_types_t::INTEGER_FORMULA,
      .criteria_        = xwpp::validation_criteria_t::NOT_BETWEEN,
      .minimum_formula_ = "=E3",
      .maximum_formula_ = "=F3",
    };
    worksheet.data_validation_cell("B5", data_validation);
  }

  // Example 3. Limiting input to an integer greater than a fixed value.
  {
    worksheet.write("A7", "Enter an integer greater than 0");

    const xwpp::data_validation_t data_validation{
      .validate_     = xwpp::validation_types_t::INTEGER,
      .criteria_     = xwpp::validation_criteria_t::GREATER_THAN,
      .value_number_ = 0,
    };
    worksheet.data_validation_cell("B7", data_validation);
  }

  // Example 4. Limiting input to an integer less than a fixed value.
  {
    worksheet.write("A9", "Enter an integer less than 10");

    const xwpp::data_validation_t data_validation{
      .validate_     = xwpp::validation_types_t::INTEGER,
      .criteria_     = xwpp::validation_criteria_t::LESS_THAN,
      .value_number_ = 10,
    };
    worksheet.data_validation_cell("B9", data_validation);
  }

  // Example 5. Limiting input to a decimal in a fixed range.
  {
    worksheet.write("A11", "Enter a decimal between 0.1 and 0.5");

    const xwpp::data_validation_t data_validation{
      .validate_       = xwpp::validation_types_t::DECIMAL,
      .criteria_       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number_ = 0.1,
      .maximum_number_ = 0.5,
    };
    worksheet.data_validation_cell("B11", data_validation);
  }

  // Example 6. Limiting input to a value in a dropdown list.
  {
    worksheet.write("A13", "Select a value from a dropdown list");

    const std::vector<std::string> list{"open", "high", "close"};

    const xwpp::data_validation_t data_validation{
      .validate_   = xwpp::validation_types_t::LIST,
      .value_list_ = list,
    };
    worksheet.data_validation_cell("B13", data_validation);
  }

  // Example 7. Limiting input to a value in a dropdown list.
  {
    worksheet.write("A15", "Select a value from a dropdown list (using a cell range)");

    const xwpp::data_validation_t data_validation{
      .validate_      = xwpp::validation_types_t::LIST_FORMULA,
      .value_formula_ = "=$E$4:$G$4",
    };
    worksheet.data_validation_cell("B15", data_validation);
  }

  // Example 8. Limiting input to a date in a fixed range.
  {
    worksheet.write("A17", "Enter a date between 1/1/2024 and 12/12/2024");

    const xwpp::data_validation_t data_validation{
      .validate_         = xwpp::validation_types_t::DATE,
      .criteria_         = xwpp::validation_criteria_t::BETWEEN,
      .minimum_datetime_ = {.year_ = 2024, .month_ = 1,  .day_ = 1,  .hour_ = 0, .min_ = 0, .sec_ = 0.},
      .maximum_datetime_ = {.year_ = 2024, .month_ = 12, .day_ = 12, .hour_ = 0, .min_ = 0, .sec_ = 0.},
    };
    worksheet.data_validation_cell("B17", data_validation);
  }

  // Example 9. Limiting input to a time in a fixed range.
  {
    worksheet.write("A19", "Enter a time between 6:00 and 12:00");

    const xwpp::data_validation_t data_validation{
      .validate_         = xwpp::validation_types_t::TIME,
      .criteria_         = xwpp::validation_criteria_t::BETWEEN,
      .minimum_datetime_ = {.year_ = 0, .month_ = 0, .day_ = 0, .hour_ = 6,  .min_ = 0, .sec_ = 0.},
      .maximum_datetime_ = {.year_ = 0, .month_ = 0, .day_ = 0, .hour_ = 12, .min_ = 0, .sec_ = 0.},
    };
    worksheet.data_validation_cell("B19", data_validation);
  }

  // Example 10. Limiting input to a string greater than a fixed length.
  {
    worksheet.write("A21", "Enter a string longer than 3 characters");

    const xwpp::data_validation_t data_validation{
      .validate_     = xwpp::validation_types_t::LENGTH,
      .criteria_     = xwpp::validation_criteria_t::GREATER_THAN,
      .value_number_ = 3,
    };
    worksheet.data_validation_cell("B21", data_validation);
  }

  // Example 11. Limiting input based on a formula.
  {
    worksheet.write("A23", "Enter a value if the following is true \"=AND(F5=50,G5=60)\"");

    const xwpp::data_validation_t data_validation{
      .validate_      = xwpp::validation_types_t::CUSTOM_FORMULA,
      .value_formula_ = "=AND(F5=50,G5=60)",
    };
    worksheet.data_validation_cell("B23", data_validation);
  }

  // Example 12. Displaying and modifying data validation messages.
  {
    worksheet.write("A25", "Displays a message when you select the cell");

    const xwpp::data_validation_t data_validation{
      .validate_       = xwpp::validation_types_t::INTEGER,
      .criteria_       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number_ = 1,
      .maximum_number_ = 100,
      .input_title_    = "Enter an integer:",
      .input_message_  = "between 1 and 100",
    };
    worksheet.data_validation_cell("B25", data_validation);
  }

  // Example 13. Displaying and modifying data validation messages.
  {
    worksheet.write("A27", "Display a custom error message when integer isn't between 1 and 100");

    const xwpp::data_validation_t data_validation{
      .validate_       = xwpp::validation_types_t::INTEGER,
      .criteria_       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number_ = 1,
      .maximum_number_ = 100,
      .input_title_    = "Enter an integer:",
      .input_message_  = "between 1 and 100",
      .error_title_    = "Input value is not valid!",
      .error_message_  = "It should be an integer between 1 and 100",
    };
    worksheet.data_validation_cell("B27", data_validation);
  }

  // Example 14. Displaying and modifying data validation messages.
  {
    worksheet.write("A29", "Display a custom info message when integer isn't between 1 and 100");

    const xwpp::data_validation_t data_validation{
      .validate_       = xwpp::validation_types_t::INTEGER,
      .criteria_       = xwpp::validation_criteria_t::BETWEEN,
      .error_type_     = xwpp::validation_error_types_t::INFORMATION,
      .minimum_number_ = 1,
      .maximum_number_ = 100,
      .input_title_    = "Enter an integer:",
      .input_message_  = "between 1 and 100",
      .error_title_    = "Input value is not valid!",
      .error_message_  = "It should be an integer between 1 and 100",
    };
    worksheet.data_validation_cell("B29", data_validation);
  }

  workbook.save("data_validate.xlsx");
}
