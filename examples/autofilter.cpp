/*
 * Example of adding an autofilter to a worksheet in Excel using
 * Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

#include <string>
#include <vector>

namespace
{

void write_worksheet_header(xwpp::worksheet_t& worksheet, const xwpp::format_t* header)
{
  // Make the columns wider for clarity.
  worksheet.set_column(0, 3, 12);

  // Write the column headers.
  worksheet.set_row(0, 20, header);
  worksheet.write(0, 0, "Region");
  worksheet.write(0, 1, "Item");
  worksheet.write(0, 2, "Volume");
  worksheet.write(0, 3, "Month");
}

}

int main()
{
  struct row
  {
    std::string region;
    std::string item;
    int volume{0};
    std::string month;
  };

  std::vector<row> data{
    {.region = "East",  .item = "Apple",  .volume = 9000,  .month = "July"     },
    {.region = "East",  .item = "Apple",  .volume = 5000,  .month = "July"     },
    {.region = "South", .item = "Orange", .volume = 9000,  .month = "September"},
    {.region = "North", .item = "Apple",  .volume = 2000,  .month = "November" },
    {.region = "West",  .item = "Apple",  .volume = 9000,  .month = "November" },
    {.region = "South", .item = "Pear",   .volume = 7000,  .month = "October"  },
    {.region = "North", .item = "Pear",   .volume = 9000,  .month = "August"   },
    {.region = "West",  .item = "Orange", .volume = 1000,  .month = "December" },
    {.region = "West",  .item = "Grape",  .volume = 1000,  .month = "November" },
    {.region = "South", .item = "Pear",   .volume = 10000, .month = "April"    },
    {.region = "West",  .item = "Grape",  .volume = 6000,  .month = "January"  },
    {.region = "South", .item = "Orange", .volume = 3000,  .month = "May"      },
    {.region = "North", .item = "Apple",  .volume = 3000,  .month = "December" },
    {.region = "South", .item = "Apple",  .volume = 7000,  .month = "February" },
    {.region = "West",  .item = "Grape",  .volume = 1000,  .month = "December" },
    {.region = "East",  .item = "Grape",  .volume = 8000,  .month = "February" },
    {.region = "South", .item = "Grape",  .volume = 10000, .month = "June"     },
    {.region = "West",  .item = "Pear",   .volume = 7000,  .month = "December" },
    {.region = "South", .item = "Apple",  .volume = 2000,  .month = "October"  },
    {.region = "East",  .item = "Grape",  .volume = 7000,  .month = "December" },
    {.region = "North", .item = "Grape",  .volume = 6000,  .month = "April"    },
    {.region = "East",  .item = "Pear",   .volume = 8000,  .month = "February" },
    {.region = "North", .item = "Apple",  .volume = 7000,  .month = "August"   },
    {.region = "North", .item = "Orange", .volume = 7000,  .month = "July"     },
    {.region = "North", .item = "Apple",  .volume = 6000,  .month = "June"     },
    {.region = "South", .item = "Grape",  .volume = 8000,  .month = "September"},
    {.region = "West",  .item = "Apple",  .volume = 3000,  .month = "October"  },
    {.region = "South", .item = "Orange", .volume = 10000, .month = "November" },
    {.region = "West",  .item = "Grape",  .volume = 4000,  .month = "July"     },
    {.region = "North", .item = "Orange", .volume = 5000,  .month = "August"   },
    {.region = "East",  .item = "Orange", .volume = 1000,  .month = "November" },
    {.region = "East",  .item = "Orange", .volume = 4000,  .month = "October"  },
    {.region = "North", .item = "Grape",  .volume = 5000,  .month = "August"   },
    {.region = "East",  .item = "Apple",  .volume = 1000,  .month = "December" },
    {.region = "South", .item = "Apple",  .volume = 10000, .month = "March"    },
    {.region = "East",  .item = "Grape",  .volume = 7000,  .month = "October"  },
    {.region = "West",  .item = "Grape",  .volume = 1000,  .month = "September"},
    {.region = "East",  .item = "Grape",  .volume = 10000, .month = "October"  },
    {.region = "South", .item = "Orange", .volume = 8000,  .month = "March"    },
    {.region = "North", .item = "Apple",  .volume = 4000,  .month = "July"     },
    {.region = "South", .item = "Orange", .volume = 5000,  .month = "July"     },
    {.region = "West",  .item = "Apple",  .volume = 4000,  .month = "June"     },
    {.region = "East",  .item = "Apple",  .volume = 5000,  .month = "April"    },
    {.region = "North", .item = "Pear",   .volume = 3000,  .month = "August"   },
    {.region = "East",  .item = "Grape",  .volume = 9000,  .month = "November" },
    {.region = "North", .item = "Orange", .volume = 8000,  .month = "October"  },
    {.region = "East",  .item = "Apple",  .volume = 10000, .month = "June"     },
    {.region = "South", .item = "Pear",   .volume = 1000,  .month = "December" },
    {.region = "North", .item = "Grape",  .volume = 10000, .month = "July"     },
    {.region = "East",  .item = "Grape",  .volume = 6000,  .month = "February" }
  };

  const xwpp::row_col_options_t hidden{.hidden = true};

  xwpp::workbook_t workbook;
  const xwpp::format_t* header = workbook.format_builder().bold().build();

  // Example 1. Autofilter without conditions.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Write the row data.
    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);
  }

  // Example 2. Autofilter with a filter condition in the first column.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Write the row data.
    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      // It isn't sufficient to just apply the filter condition below. We
      // must also hide the rows that don't match the criteria since Excel
      // doesn't do that automatically.
      if(region == "East")
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // Hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }

      // Note, the if() statement above is written to match the logic of the
      // criteria in worksheet::filter_column() below. However you could get
      // the same results with the following simpler, but reversed, code:
      //
      //     if (region != "East")
      //     {
      //       worksheet.set_row(i + 1, hidden);
      //     }
      //
      // The same applies to the Examples 3-6 as well.
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const xwpp::filter_rule_t filter_rule2{
      .criteria     = xwpp::filter_criteria_t::EQUAL_TO,
      .value_string = "East",
    };
    worksheet.filter_column(0, filter_rule2);
  }

  // Example 3. Autofilter with a dual filter condition in one of the columns.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Write the row data.
    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      if(region == "East" || region == "South")
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // We need to hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const xwpp::filter_rule_t filter_rule3a{
      .criteria     = xwpp::filter_criteria_t::EQUAL_TO,
      .value_string = "East",
    };
    const xwpp::filter_rule_t filter_rule3b{
      .criteria     = xwpp::filter_criteria_t::EQUAL_TO,
      .value_string = "South",
    };
    worksheet.filter_column2(0, filter_rule3a, filter_rule3b, xwpp::filter_operator_t::OR);
  }

  // Example 4. Autofilter with filter conditions in two columns.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Write the row data.
    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      if(region == "East" && volume > 3000 && volume < 8000)
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // We need to hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const xwpp::filter_rule_t filter_rule4a{
      .criteria     = xwpp::filter_criteria_t::EQUAL_TO,
      .value_string = "East",
    };
    const xwpp::filter_rule_t filter_rule4b{
      .criteria = xwpp::filter_criteria_t::GREATER_THAN,
      .value    = 3000,
    };
    const xwpp::filter_rule_t filter_rule4c{
      .criteria = xwpp::filter_criteria_t::LESS_THAN,
      .value    = 8000,
    };
    worksheet.filter_column(0, filter_rule4a);
    worksheet.filter_column2(2, filter_rule4b, filter_rule4c, xwpp::filter_operator_t::AND);
  }

  // Example 5. Autofilter with a dual filter condition in one of the columns.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Write the row data.
    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      if(region == "East" || region == "North" || region == "South")
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // We need to hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const std::vector<std::string> list{"East", "North", "South"};
    worksheet.filter_list(0, list);
  }

  // Example 6. Autofilter with filter for blanks.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    // Simulate one blank cell in the data, to test the filter.
    data[5].region.clear();

    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      if(region.empty())
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // We need to hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const xwpp::filter_rule_t filter_rule6{.criteria = xwpp::filter_criteria_t::BLANKS};
    worksheet.filter_column(0, filter_rule6);
  }

  // Example 7. Autofilter with filter for non-blanks.
  // Set up the worksheet data.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
    write_worksheet_header(worksheet, header);

    for(xwpp::row_num_t i{0}; const auto& [region, item, volume, month]: data)
    {
      worksheet.write(i + 1, 0, region);
      worksheet.write(i + 1, 1, item);
      worksheet.write(i + 1, 2, volume);
      worksheet.write(i + 1, 3, month);

      if(!region.empty())
      {
        // Row matches the filter, no further action required.
      }
      else
      {
        // We need to hide rows that don't match the filter.
        worksheet.set_row(i + 1, hidden);
      }
      i++;
    }

    // Add the autofilter.
    worksheet.autofilter(0, 0, 50, 3);

    // Add the filter criteria.
    const xwpp::filter_rule_t filter_rule7{.criteria = xwpp::filter_criteria_t::NON_BLANKS};
    worksheet.filter_column(0, filter_rule7);
  }

  workbook.save("autofilter.xlsx");
}
