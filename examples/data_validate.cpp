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
  const xwpp::format_t* format = workbook.format_builder()
                                   .border(xwpp::format_borders_t::THIN)
                                   .fg_color(xwpp::color_t{0xC6EFCE})
                                   .bold()
                                   .text_wrap()
                                   .align(xwpp::format_vertical_alignments_t::CENTER)
                                   .indent(1)
                                   .build();

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
      .validate       = xwpp::validation_types_t::INTEGER,
      .criteria       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number = 1,
      .maximum_number = 10,
    };
    worksheet.data_validation_cell("B3", data_validation);
  }

  // Example 2. Limiting input to an integer outside a fixed range.
  {
    worksheet.write("A5", "Enter an integer that is not between 1 and 10 (using cell references)");

    const xwpp::data_validation_t data_validation{
      .validate        = xwpp::validation_types_t::INTEGER_FORMULA,
      .criteria        = xwpp::validation_criteria_t::NOT_BETWEEN,
      .minimum_formula = "=E3",
      .maximum_formula = "=F3",
    };
    worksheet.data_validation_cell("B5", data_validation);
  }

  // Example 3. Limiting input to an integer greater than a fixed value.
  {
    worksheet.write("A7", "Enter an integer greater than 0");

    const xwpp::data_validation_t data_validation{
      .validate     = xwpp::validation_types_t::INTEGER,
      .criteria     = xwpp::validation_criteria_t::GREATER_THAN,
      .value_number = 0,
    };
    worksheet.data_validation_cell("B7", data_validation);
  }

  // Example 4. Limiting input to an integer less than a fixed value.
  {
    worksheet.write("A9", "Enter an integer less than 10");

    const xwpp::data_validation_t data_validation{
      .validate     = xwpp::validation_types_t::INTEGER,
      .criteria     = xwpp::validation_criteria_t::LESS_THAN,
      .value_number = 10,
    };
    worksheet.data_validation_cell("B9", data_validation);
  }

  // Example 5. Limiting input to a decimal in a fixed range.
  {
    worksheet.write("A11", "Enter a decimal between 0.1 and 0.5");

    const xwpp::data_validation_t data_validation{
      .validate       = xwpp::validation_types_t::DECIMAL,
      .criteria       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number = 0.1,
      .maximum_number = 0.5,
    };
    worksheet.data_validation_cell("B11", data_validation);
  }

  // Example 6. Limiting input to a value in a dropdown list.
  {
    worksheet.write("A13", "Select a value from a dropdown list");

    const std::vector<std::string> list{"open", "high", "close"};

    const xwpp::data_validation_t data_validation{
      .validate   = xwpp::validation_types_t::LIST,
      .value_list = list,
    };
    worksheet.data_validation_cell("B13", data_validation);
  }

  // Example 7. Limiting input to a value in a dropdown list.
  {
    worksheet.write("A15", "Select a value from a dropdown list (using a cell range)");

    const xwpp::data_validation_t data_validation{
      .validate      = xwpp::validation_types_t::LIST_FORMULA,
      .value_formula = "=$E$4:$G$4",
    };
    worksheet.data_validation_cell("B15", data_validation);
  }

  // Example 8. Limiting input to a date in a fixed range.
  {
    worksheet.write("A17", "Enter a date between 1/1/2024 and 12/12/2024");

    const xwpp::data_validation_t data_validation{
      .validate         = xwpp::validation_types_t::DATE,
      .criteria         = xwpp::validation_criteria_t::BETWEEN,
      .minimum_datetime = {.year = 2024, .month = 1,  .day = 1,  .hour = 0, .min = 0, .sec = 0.},
      .maximum_datetime = {.year = 2024, .month = 12, .day = 12, .hour = 0, .min = 0, .sec = 0.},
    };
    worksheet.data_validation_cell("B17", data_validation);
  }

  // Example 9. Limiting input to a time in a fixed range.
  {
    worksheet.write("A19", "Enter a time between 6:00 and 12:00");

    const xwpp::data_validation_t data_validation{
      .validate         = xwpp::validation_types_t::TIME,
      .criteria         = xwpp::validation_criteria_t::BETWEEN,
      .minimum_datetime = {.year = 0, .month = 0, .day = 0, .hour = 6,  .min = 0, .sec = 0.},
      .maximum_datetime = {.year = 0, .month = 0, .day = 0, .hour = 12, .min = 0, .sec = 0.},
    };
    worksheet.data_validation_cell("B19", data_validation);
  }

  // Example 10. Limiting input to a string greater than a fixed length.
  {
    worksheet.write("A21", "Enter a string longer than 3 characters");

    const xwpp::data_validation_t data_validation{
      .validate     = xwpp::validation_types_t::LENGTH,
      .criteria     = xwpp::validation_criteria_t::GREATER_THAN,
      .value_number = 3,
    };
    worksheet.data_validation_cell("B21", data_validation);
  }

  // Example 11. Limiting input based on a formula.
  {
    worksheet.write("A23", "Enter a value if the following is true \"=AND(F5=50,G5=60)\"");

    const xwpp::data_validation_t data_validation{
      .validate      = xwpp::validation_types_t::CUSTOM_FORMULA,
      .value_formula = "=AND(F5=50,G5=60)",
    };
    worksheet.data_validation_cell("B23", data_validation);
  }

  // Example 12. Displaying and modifying data validation messages.
  {
    worksheet.write("A25", "Displays a message when you select the cell");

    const xwpp::data_validation_t data_validation{
      .validate       = xwpp::validation_types_t::INTEGER,
      .criteria       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number = 1,
      .maximum_number = 100,
      .input_title    = "Enter an integer:",
      .input_message  = "between 1 and 100",
    };
    worksheet.data_validation_cell("B25", data_validation);
  }

  // Example 13. Displaying and modifying data validation messages.
  {
    worksheet.write("A27", "Display a custom error message when integer isn't between 1 and 100");

    const xwpp::data_validation_t data_validation{
      .validate       = xwpp::validation_types_t::INTEGER,
      .criteria       = xwpp::validation_criteria_t::BETWEEN,
      .minimum_number = 1,
      .maximum_number = 100,
      .input_title    = "Enter an integer:",
      .input_message  = "between 1 and 100",
      .error_title    = "Input value is not valid!",
      .error_message  = "It should be an integer between 1 and 100",
    };
    worksheet.data_validation_cell("B27", data_validation);
  }

  // Example 14. Displaying and modifying data validation messages.
  {
    worksheet.write("A29", "Display a custom info message when integer isn't between 1 and 100");

    const xwpp::data_validation_t data_validation{
      .validate       = xwpp::validation_types_t::INTEGER,
      .criteria       = xwpp::validation_criteria_t::BETWEEN,
      .error_type     = xwpp::validation_error_types_t::INFORMATION,
      .minimum_number = 1,
      .maximum_number = 100,
      .input_title    = "Enter an integer:",
      .input_message  = "between 1 and 100",
      .error_title    = "Input value is not valid!",
      .error_message  = "It should be an integer between 1 and 100",
    };
    worksheet.data_validation_cell("B29", data_validation);
  }

  workbook.save("data_validate.xlsx");
}
