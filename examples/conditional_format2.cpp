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

void write_worksheet_data(xwpp::worksheet_t& worksheet)
{
  uint8_t data[10][10] = {
      {34, 72,  38, 30, 75, 48, 75, 66, 84, 86},
      {6,  24,  1,  84, 54, 62, 60, 3,  26, 59},
      {28, 79,  97, 13, 85, 93, 93, 22, 5,  14},
      {27, 71,  40, 17, 18, 79, 90, 93, 29, 47},
      {88, 25,  33, 23, 67, 1,  59, 79, 47, 36},
      {24, 100, 20, 88, 29, 33, 38, 54, 54, 88},
      {6,  57,  88, 28, 10, 26, 37, 7,  41, 48},
      {52, 78,  1,  96, 26, 45, 47, 33, 96, 36},
      {60, 54,  81, 66, 81, 90, 80, 93, 12, 55},
      {70, 5,   46, 14, 71, 19, 66, 36, 41, 21},
  };

  int row, col;
  for(row = 0; row < 10; row++)
  {
    for(col = 0; col < 10; col++)
    {
      worksheet.write_number(row + 2, col + 1, data[row][col]);
    }
  }
}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet2 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet3 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet4 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet5 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet6 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet7 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet8 = workbook.add_worksheet();
  xwpp::worksheet_t& worksheet9 = workbook.add_worksheet();

  // Add a format. Light red fill with dark red text.
  xwpp::format_t* format1 = workbook.add_format();
  format1->set_bg_color(static_cast<xwpp::color_t>(0xFFC7CE));
  format1->set_font_color(static_cast<xwpp::color_t>(0x9C0006));

  // Add a format. Green fill with dark green text.
  xwpp::format_t* format2 = workbook.add_format();
  format2->set_bg_color(static_cast<xwpp::color_t>(0xC6EFCE));
  format2->set_font_color(static_cast<xwpp::color_t>(0x006100));

  xwpp::conditional_format_t conditional_format;

  // Example 1. Conditional formatting based on simple cell based criteria.
  write_worksheet_data(worksheet1);

  worksheet1.write_string(CELL("A1"), "Cells with values >= 50 are in light red. Values < 50 are in light green.");

  conditional_format.type_     = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO;
  conditional_format.value_    = 50;
  conditional_format.format_   = format1;
  worksheet1.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format           = xwpp::conditional_format_t{};
  conditional_format.type_     = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::LESS_THAN;
  conditional_format.value_    = 50;
  conditional_format.format_   = format2;
  worksheet1.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 2. Conditional formatting based on max and min values.
  write_worksheet_data(worksheet2);

  worksheet2.write_string(CELL("A1"),
                          "Values between 30 and 70 are in light red. Values outside that range are in light green.");

  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_  = xwpp::conditional_criteria_t::BETWEEN;
  conditional_format.min_value_ = 30;
  conditional_format.max_value_ = 70;
  conditional_format.format_    = format1;
  worksheet2.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_  = xwpp::conditional_criteria_t::NOT_BETWEEN;
  conditional_format.min_value_ = 30;
  conditional_format.max_value_ = 70;
  conditional_format.format_    = format2;
  worksheet2.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 3. Conditional formatting with duplicate and unique values.
  write_worksheet_data(worksheet3);

  worksheet3.write_string(CELL("A1"), "Duplicate values are in light red. Unique values are in light green.");

  conditional_format         = xwpp::conditional_format_t{};
  conditional_format.type_   = xwpp::conditional_format_types_t::DUPLICATE;
  conditional_format.format_ = format1;
  worksheet3.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format         = xwpp::conditional_format_t{};
  conditional_format.type_   = xwpp::conditional_format_types_t::UNIQUE;
  conditional_format.format_ = format2;
  worksheet3.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 4. Conditional formatting with above and below average values.
  write_worksheet_data(worksheet4);

  worksheet4.write_string(CELL("A1"),
                          "Above average values are in light red. Below average values are in light green.");

  conditional_format           = xwpp::conditional_format_t{};
  conditional_format.type_     = xwpp::conditional_format_types_t::AVERAGE;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::AVERAGE_ABOVE;
  conditional_format.format_   = format1;
  worksheet4.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format           = xwpp::conditional_format_t{};
  conditional_format.type_     = xwpp::conditional_format_types_t::AVERAGE;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::AVERAGE_BELOW;
  conditional_format.format_   = format2;
  worksheet4.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 5. Conditional formatting with top and bottom values.
  write_worksheet_data(worksheet5);

  worksheet5.write_string(CELL("A1"), "Top 10 values are in light red. Bottom 10 values are in light green.");

  conditional_format         = xwpp::conditional_format_t{};
  conditional_format.type_   = xwpp::conditional_format_types_t::TOP;
  conditional_format.value_  = 10;
  conditional_format.format_ = format1;
  worksheet5.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format         = xwpp::conditional_format_t{};
  conditional_format.type_   = xwpp::conditional_format_types_t::BOTTOM;
  conditional_format.value_  = 10;
  conditional_format.format_ = format2;
  worksheet5.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 6. Conditional formatting with multiple ranges.
  write_worksheet_data(worksheet6);

  worksheet6.write_string(
      CELL("A1"), "Cells with values >= 50 are in light red.Values < 50 are in light green. Non-contiguous ranges.");

  conditional_format              = xwpp::conditional_format_t{};
  conditional_format.type_        = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_    = xwpp::conditional_criteria_t::GREATER_THAN_OR_EQUAL_TO;
  conditional_format.value_       = 50;
  conditional_format.format_      = format1;
  conditional_format.multi_range_ = "B3:K6 B9:K12";
  worksheet6.conditional_format_range(RANGE("B3:K12"), conditional_format);

  conditional_format              = xwpp::conditional_format_t{};
  conditional_format.type_        = xwpp::conditional_format_types_t::CELL;
  conditional_format.criteria_    = xwpp::conditional_criteria_t::LESS_THAN;
  conditional_format.value_       = 50;
  conditional_format.format_      = format2;
  conditional_format.multi_range_ = "B3:K6 B9:K12";
  worksheet6.conditional_format_range(RANGE("B3:K12"), conditional_format);

  // Example 7. Conditional formatting with 2 color scales.
  // Write the worksheet data.
  for(int i = 1; i <= 12; i++)
  {
    worksheet7.write_number(i + 1, 1, i);
    worksheet7.write_number(i + 1, 3, i);
    worksheet7.write_number(i + 1, 6, i);
    worksheet7.write_number(i + 1, 8, i);
  }
  worksheet7.write_string(CELL("A1"), "Examples of color scales with default and user colors.");

  worksheet7.write_string(CELL("B2"), "2 Color Scale");
  worksheet7.write_string(CELL("D2"), "2 Color Scale + user colors");
  worksheet7.write_string(CELL("G2"), "3 Color Scale");
  worksheet7.write_string(CELL("I2"), "3 Color Scale + user colors");

  // 2 color scale with standard colors.
  conditional_format       = xwpp::conditional_format_t{};
  conditional_format.type_ = xwpp::conditional_format_types_t::TWO_COLOR_SCALE;
  worksheet7.conditional_format_range(RANGE("B3:B14"), conditional_format);

  // 2 color scale with user defined colors.
  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::TWO_COLOR_SCALE;
  conditional_format.min_color_ = static_cast<xwpp::color_t>(0xFF0000);
  conditional_format.max_color_ = static_cast<xwpp::color_t>(0x00FF00);
  worksheet7.conditional_format_range(RANGE("D3:D14"), conditional_format);

  // 3 color scale with standard colors.
  conditional_format       = xwpp::conditional_format_t{};
  conditional_format.type_ = xwpp::conditional_format_types_t::THREE_COLOR_SCALE;
  worksheet7.conditional_format_range(RANGE("G3:G14"), conditional_format);

  // 3 color scale with user defined colors.
  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::THREE_COLOR_SCALE;
  conditional_format.min_color_ = static_cast<xwpp::color_t>(0xC5D9F1);
  conditional_format.mid_color_ = static_cast<xwpp::color_t>(0x8DB4E3);
  conditional_format.max_color_ = static_cast<xwpp::color_t>(0x538ED5);
  worksheet7.conditional_format_range(RANGE("I3:I14"), conditional_format);

  // Example 8. Conditional formatting with data bars.
  // Write the worksheet data.
  for(int i = 1; i <= 12; i++)
  {
    worksheet8.write_number(i + 1, 1, i);
    worksheet8.write_number(i + 1, 3, i);
    worksheet8.write_number(i + 1, 5, i);
    worksheet8.write_number(i + 1, 7, i);
    worksheet8.write_number(i + 1, 9, i);
  }

  int data[] = {-1, -2, -3, -2, -1, 0, 1, 2, 3, 2, 1, 0};
  for(int i = 1; i <= 12; i++)
  {
    worksheet8.write_number(i + 1, 11, data[i - 1]);
    worksheet8.write_number(i + 1, 13, data[i - 1]);
  }

  worksheet8.write_string(CELL("A1"), "Examples of data bars.");

  worksheet8.write_string(CELL("B2"), "Default data bars");
  worksheet8.write_string(CELL("D2"), "Bars only");
  worksheet8.write_string(CELL("F2"), "With user color");
  worksheet8.write_string(CELL("H2"), "Solid bars");
  worksheet8.write_string(CELL("J2"), "Right to left");
  worksheet8.write_string(CELL("L2"), "Excel 2010 style");
  worksheet8.write_string(CELL("N2"), "Negative same as positive");

  conditional_format       = xwpp::conditional_format_t{};
  conditional_format.type_ = xwpp::conditional_format_types_t::DATA_BAR;
  worksheet8.conditional_format_range(RANGE("B3:B14"), conditional_format);

  conditional_format           = xwpp::conditional_format_t{};
  conditional_format.type_     = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.bar_only_ = true;
  worksheet8.conditional_format_range(RANGE("D3:D14"), conditional_format);

  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.bar_color_ = static_cast<xwpp::color_t>(0x63C384);
  worksheet8.conditional_format_range(RANGE("F3:F14"), conditional_format);

  conditional_format            = xwpp::conditional_format_t{};
  conditional_format.type_      = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.bar_solid_ = true;
  worksheet8.conditional_format_range(RANGE("H3:H14"), conditional_format);

  conditional_format                = xwpp::conditional_format_t{};
  conditional_format.type_          = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.bar_direction_ = xwpp::conditional_format_bar_direction_t::RIGHT_TO_LEFT;
  worksheet8.conditional_format_range(RANGE("J3:J14"), conditional_format);

  conditional_format                = xwpp::conditional_format_t{};
  conditional_format.type_          = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.data_bar_2010_ = true;
  worksheet8.conditional_format_range(RANGE("L3:L14"), conditional_format);

  conditional_format                                 = xwpp::conditional_format_t{};
  conditional_format.type_                           = xwpp::conditional_format_types_t::DATA_BAR;
  conditional_format.bar_negative_color_same_        = true;
  conditional_format.bar_negative_border_color_same_ = true;
  worksheet8.conditional_format_range(RANGE("N3:N14"), conditional_format);
  conditional_format = xwpp::conditional_format_t{};

  // Example 9. Conditional formatting with icon sets.
  // Write the worksheet data.
  for(int i = 1; i <= 3; i++)
  {
    worksheet9.write_number(2, i, i);
    worksheet9.write_number(3, i, i);
    worksheet9.write_number(4, i, i);
    worksheet9.write_number(5, i, i);
  }

  for(int i = 1; i <= 4; i++)
  {
    worksheet9.write_number(6, i, i);
  }

  for(int i = 1; i <= 5; i++)
  {
    worksheet9.write_number(7, i, i);
    worksheet9.write_number(8, i, i);
  }

  worksheet9.write_string(CELL("A1"), "Examples of conditional formats with icon sets.");

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED;
  worksheet9.conditional_format_range(RANGE("B3:D3"), conditional_format);

  conditional_format                = xwpp::conditional_format_t{};
  conditional_format.type_          = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_    = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED;
  conditional_format.reverse_icons_ = true;
  worksheet9.conditional_format_range(RANGE("B4:D4"), conditional_format);

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_TRAFFIC_LIGHTS_UNRIMMED;
  conditional_format.icons_only_ = true;
  worksheet9.conditional_format_range(RANGE("B5:D5"), conditional_format);

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::THREE_ARROWS_COLORED;
  worksheet9.conditional_format_range(RANGE("B6:D6"), conditional_format);

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FOUR_ARROWS_COLORED;
  worksheet9.conditional_format_range(RANGE("B7:E7"), conditional_format);

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FIVE_ARROWS_COLORED;
  worksheet9.conditional_format_range(RANGE("B8:F8"), conditional_format);

  conditional_format             = xwpp::conditional_format_t{};
  conditional_format.type_       = xwpp::conditional_format_types_t::ICON_SETS;
  conditional_format.icon_style_ = xwpp::conditional_icon_types_t::FIVE_RATINGS;
  worksheet9.conditional_format_range(RANGE("B9:F9"), conditional_format);

  workbook.save("conditional_format.xlsx");
}
