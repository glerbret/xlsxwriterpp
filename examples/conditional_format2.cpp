/*
 * An example of how to add conditional formatting to an Xlsxwriter++ file.
 *
 * Conditional formatting allows you to apply a format to a cell or a
 * range of cells based on certain criteria.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <array>
#include <vector>

namespace
{

void write_worksheet_data(xwpp::worksheet_t& worksheet)
{
  const std::array<std::array<uint8_t, 10>, 10> data{
    {
     {34, 72, 38, 30, 75, 48, 75, 66, 84, 86},
     {6, 24, 1, 84, 54, 62, 60, 3, 26, 59},
     {28, 79, 97, 13, 85, 93, 93, 22, 5, 14},
     {27, 71, 40, 17, 18, 79, 90, 93, 29, 47},
     {88, 25, 33, 23, 67, 1, 59, 79, 47, 36},
     {24, 100, 20, 88, 29, 33, 38, 54, 54, 88},
     {6, 57, 88, 28, 10, 26, 37, 7, 41, 48},
     {52, 78, 1, 96, 26, 45, 47, 33, 96, 36},
     {60, 54, 81, 66, 81, 90, 80, 93, 12, 55},
     {70, 5, 46, 14, 71, 19, 66, 36, 41, 21},
     }
  };

  for(xwpp::row_num_t row_num{0}; const auto& row: data)
  {
    for(xwpp::col_num_t col_num{0}; const auto value: row)
    {
      worksheet.write_number(row_num, col_num, value);
      col_num++;
    }
    row_num++;
  }
}

}

int main()
{
  xwpp::workbook_t workbook;

  // Add a format. Light red fill with dark red text.
  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bg_color(static_cast<xwpp::color_t>(0xFFC7CE));
  format1->set_font_color(static_cast<xwpp::color_t>(0x9C0006));

  // Add a format. Green fill with dark green text.
  xwpp::format_t* format2 = workbook.add_format();
  format2->set_bg_color(static_cast<xwpp::color_t>(0xC6EFCE));
  format2->set_font_color(static_cast<xwpp::color_t>(0x006100));

  // Example 1. Conditional formatting based on simple cell based criteria.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(CELL("A1"), "Cells with values >= 50 are in light red. Values < 50 are in light green.");

    const xwpp::conditional_format_t conditional_format1{
      .type_     = xwpp::conditional_format_types_t::CELL,
      .criteria_ = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO,
      .value_    = 50,
      .format_   = format1,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_     = xwpp::conditional_format_types_t::CELL,
      .criteria_ = xwpp::conditional_criteria_t::LESS_THAN,
      .value_    = 50,
      .format_   = format2,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 2. Conditional formatting based on max and min values.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(CELL("A1"),
                           "Values between 30 and 70 are in light red. Values outside that range are in light green.");

    const xwpp::conditional_format_t conditional_format1{
      .type_      = xwpp::conditional_format_types_t::CELL,
      .criteria_  = xwpp::conditional_criteria_t::BETWEEN,
      .format_    = format1,
      .min_value_ = 30,
      .max_value_ = 70,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_      = xwpp::conditional_format_types_t::CELL,
      .criteria_  = xwpp::conditional_criteria_t::NOT_BETWEEN,
      .format_    = format2,
      .min_value_ = 30,
      .max_value_ = 70,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 3. Conditional formatting with duplicate and unique values.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(CELL("A1"), "Duplicate values are in light red. Unique values are in light green.");

    const xwpp::conditional_format_t conditional_format1{
      .type_   = xwpp::conditional_format_types_t::DUPLICATE,
      .format_ = format1,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_   = xwpp::conditional_format_types_t::UNIQUE,
      .format_ = format2,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 4. Conditional formatting with above and below average values.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(CELL("A1"),
                           "Above average values are in light red. Below average values are in light green.");

    const xwpp::conditional_format_t conditional_format1{
      .type_     = xwpp::conditional_format_types_t::AVERAGE,
      .criteria_ = xwpp::conditional_criteria_t::AVERAGE_ABOVE,
      .format_   = format1,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_     = xwpp::conditional_format_types_t::AVERAGE,
      .criteria_ = xwpp::conditional_criteria_t::AVERAGE_BELOW,
      .format_   = format2,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 5. Conditional formatting with top and bottom values.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(CELL("A1"), "Top 10 values are in light red. Bottom 10 values are in light green.");

    const xwpp::conditional_format_t conditional_format1{
      .type_   = xwpp::conditional_format_types_t::TOP,
      .value_  = 10,
      .format_ = format1,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_   = xwpp::conditional_format_types_t::BOTTOM,
      .value_  = 10,
      .format_ = format2,
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 6. Conditional formatting with multiple ranges.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    write_worksheet_data(worksheet);

    worksheet.write_string(
      CELL("A1"), "Cells with values >= 50 are in light red.Values < 50 are in light green. Non-contiguous ranges.");

    const xwpp::conditional_format_t conditional_format1{
      .type_        = xwpp::conditional_format_types_t::CELL,
      .criteria_    = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO,
      .value_       = 50,
      .format_      = format1,
      .multi_range_ = "B3:K6 B9:K12",
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_        = xwpp::conditional_format_types_t::CELL,
      .criteria_    = xwpp::conditional_criteria_t::LESS_THAN,
      .value_       = 50,
      .format_      = format2,
      .multi_range_ = "B3:K6 B9:K12",
    };
    worksheet.conditional_format_range(RANGE("B3:K12"), conditional_format2);
  }

  // Example 7. Conditional formatting with 2 color scales.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    for(int i{1}; i <= 12; i++)
    {
      worksheet.write_number(i + 1, 1, i);
      worksheet.write_number(i + 1, 3, i);
      worksheet.write_number(i + 1, 6, i);
      worksheet.write_number(i + 1, 8, i);
    }
    worksheet.write_string(CELL("A1"), "Examples of color scales with default and user colors.");

    worksheet.write_string(CELL("B2"), "2 Color Scale");
    worksheet.write_string(CELL("D2"), "2 Color Scale + user colors");
    worksheet.write_string(CELL("G2"), "3 Color Scale");
    worksheet.write_string(CELL("I2"), "3 Color Scale + user colors");

    // 2 color scale with standard colors.
    const xwpp::conditional_format_t conditional_format1{
      .type_ = xwpp::conditional_format_types_t::TWO_COLOR_SCALE,
    };
    worksheet.conditional_format_range(RANGE("B3:B14"), conditional_format1);

    // 2 color scale with user defined colors.
    const xwpp::conditional_format_t conditional_format2{
      .type_      = xwpp::conditional_format_types_t::TWO_COLOR_SCALE,
      .min_color_ = static_cast<xwpp::color_t>(0xFF0000),
      .max_color_ = static_cast<xwpp::color_t>(0x00FF00),
    };
    worksheet.conditional_format_range(RANGE("D3:D14"), conditional_format2);

    // 3 color scale with standard colors.
    const xwpp::conditional_format_t conditional_format3{
      .type_ = xwpp::conditional_format_types_t::THREE_COLOR_SCALE,
    };
    worksheet.conditional_format_range(RANGE("G3:G14"), conditional_format3);

    // 3 color scale with user defined colors.
    const xwpp::conditional_format_t conditional_format4{
      .type_      = xwpp::conditional_format_types_t::THREE_COLOR_SCALE,
      .min_color_ = static_cast<xwpp::color_t>(0xC5D9F1),
      .mid_color_ = static_cast<xwpp::color_t>(0x8DB4E3),
      .max_color_ = static_cast<xwpp::color_t>(0x538ED5),
    };
    worksheet.conditional_format_range(RANGE("I3:I14"), conditional_format4);
  }

  // Example 8. Conditional formatting with data bars.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    for(int i{1}; i <= 12; i++)
    {
      worksheet.write_number(i + 1, 1, i);
      worksheet.write_number(i + 1, 3, i);
      worksheet.write_number(i + 1, 5, i);
      worksheet.write_number(i + 1, 7, i);
      worksheet.write_number(i + 1, 9, i);
    }

    const std::vector<int> data{-1, -2, -3, -2, -1, 0, 1, 2, 3, 2, 1, 0};
    for(xwpp::row_num_t row_num{2}; const auto value: data)
    {
      worksheet.write_number(row_num, 11, value);
      worksheet.write_number(row_num, 13, value);
      row_num++;
    }

    worksheet.write_string(CELL("A1"), "Examples of data bars.");

    worksheet.write_string(CELL("B2"), "Default data bars");
    worksheet.write_string(CELL("D2"), "Bars only");
    worksheet.write_string(CELL("F2"), "With user color");
    worksheet.write_string(CELL("H2"), "Solid bars");
    worksheet.write_string(CELL("J2"), "Right to left");
    worksheet.write_string(CELL("L2"), "Excel 2010 style");
    worksheet.write_string(CELL("N2"), "Negative same as positive");

    const xwpp::conditional_format_t conditional_format1{
      .type_ = xwpp::conditional_format_types_t::DATA_BAR,
    };
    worksheet.conditional_format_range(RANGE("B3:B14"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_     = xwpp::conditional_format_types_t::DATA_BAR,
      .bar_only_ = true,
    };
    worksheet.conditional_format_range(RANGE("D3:D14"), conditional_format2);

    const xwpp::conditional_format_t conditional_format3{
      .type_      = xwpp::conditional_format_types_t::DATA_BAR,
      .bar_color_ = static_cast<xwpp::color_t>(0x63C384),
    };
    worksheet.conditional_format_range(RANGE("F3:F14"), conditional_format3);

    const xwpp::conditional_format_t conditional_format4{
      .type_      = xwpp::conditional_format_types_t::DATA_BAR,
      .bar_solid_ = true,
    };
    worksheet.conditional_format_range(RANGE("H3:H14"), conditional_format4);

    const xwpp::conditional_format_t conditional_format5{
      .type_          = xwpp::conditional_format_types_t::DATA_BAR,
      .bar_direction_ = xwpp::conditional_format_bar_direction_t::RIGHT_TO_LEFT,
    };
    worksheet.conditional_format_range(RANGE("J3:J14"), conditional_format5);

    const xwpp::conditional_format_t conditional_format6{
      .type_          = xwpp::conditional_format_types_t::DATA_BAR,
      .data_bar_2010_ = true,
    };
    worksheet.conditional_format_range(RANGE("L3:L14"), conditional_format6);

    const xwpp::conditional_format_t conditional_format7{
      .type_                           = xwpp::conditional_format_types_t::DATA_BAR,
      .bar_negative_color_same_        = true,
      .bar_negative_border_color_same_ = true,
    };
    worksheet.conditional_format_range(RANGE("N3:N14"), conditional_format7);
  }

  // Example 9. Conditional formatting with icon sets.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();

    // Write the worksheet data.
    for(int i{1}; i <= 3; i++)
    {
      worksheet.write_number(2, static_cast<xwpp::col_num_t>(i), i);
      worksheet.write_number(3, static_cast<xwpp::col_num_t>(i), i);
      worksheet.write_number(4, static_cast<xwpp::col_num_t>(i), i);
      worksheet.write_number(5, static_cast<xwpp::col_num_t>(i), i);
    }

    for(int i{1}; i <= 4; i++)
    {
      worksheet.write_number(6, static_cast<xwpp::col_num_t>(i), i);
    }

    for(int i{1}; i <= 5; i++)
    {
      worksheet.write_number(7, static_cast<xwpp::col_num_t>(i), i);
      worksheet.write_number(8, static_cast<xwpp::col_num_t>(i), i);
    }

    worksheet.write_string(CELL("A1"), "Examples of conditional formats with icon sets.");

    const xwpp::conditional_format_t conditional_format1{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED,
    };
    worksheet.conditional_format_range(RANGE("B3:D3"), conditional_format1);

    const xwpp::conditional_format_t conditional_format2{
      .type_          = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_    = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED,
      .reverse_icons_ = true,
    };
    worksheet.conditional_format_range(RANGE("B4:D4"), conditional_format2);

    const xwpp::conditional_format_t conditional_format3{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED,
      .icons_only_ = true,
    };
    worksheet.conditional_format_range(RANGE("B5:D5"), conditional_format3);

    const xwpp::conditional_format_t conditional_format4{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::THREE_ARROWS_COLORED,
    };
    worksheet.conditional_format_range(RANGE("B6:D6"), conditional_format4);

    const xwpp::conditional_format_t conditional_format5{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::FOUR_ARROWS_COLORED,
    };
    worksheet.conditional_format_range(RANGE("B7:E7"), conditional_format5);

    const xwpp::conditional_format_t conditional_format6{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::FIVE_ARROWS_COLORED,
    };
    worksheet.conditional_format_range(RANGE("B8:F8"), conditional_format6);

    const xwpp::conditional_format_t conditional_format7{
      .type_       = xwpp::conditional_format_types_t::ICON_SETS,
      .icon_style_ = xwpp::conditional_icon_types_t::FIVE_RATINGS,
    };
    worksheet.conditional_format_range(RANGE("B9:F9"), conditional_format7);
  }

  workbook.save("conditional_format.xlsx");
}
