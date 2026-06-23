/*
 * Example of adding an autofilter to a worksheet in Excel using
 * Xlsxwriter++.
 *
 * Copyright 2026, Grégory Lerbret
 *
 */

#include "xlsxwriterpp.h"

void write_worksheet_header(xwpp::worksheet_t& worksheet, const xwpp::format_t* header)
{
  // Make the columns wider for clarity.
  worksheet.set_column(0, 3, 12);

  // Write the column headers.
  worksheet.set_row(0, 20, header, std::nullopt);
  worksheet.write_string(0, 0, "Region");
  worksheet.write_string(0, 1, "Item");
  worksheet.write_string(0, 2, "Volume");
  worksheet.write_string(0, 3, "Month");
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

  struct row
  {
    std::string region;
    std::string item;
    int volume;
    std::string month;
  };

  std::vector<row> data = {
      {"East",  "Apple",  9000,  "July"     },
      {"East",  "Apple",  5000,  "July"     },
      {"South", "Orange", 9000,  "September"},
      {"North", "Apple",  2000,  "November" },
      {"West",  "Apple",  9000,  "November" },
      {"South", "Pear",   7000,  "October"  },
      {"North", "Pear",   9000,  "August"   },
      {"West",  "Orange", 1000,  "December" },
      {"West",  "Grape",  1000,  "November" },
      {"South", "Pear",   10000, "April"    },
      {"West",  "Grape",  6000,  "January"  },
      {"South", "Orange", 3000,  "May"      },
      {"North", "Apple",  3000,  "December" },
      {"South", "Apple",  7000,  "February" },
      {"West",  "Grape",  1000,  "December" },
      {"East",  "Grape",  8000,  "February" },
      {"South", "Grape",  10000, "June"     },
      {"West",  "Pear",   7000,  "December" },
      {"South", "Apple",  2000,  "October"  },
      {"East",  "Grape",  7000,  "December" },
      {"North", "Grape",  6000,  "April"    },
      {"East",  "Pear",   8000,  "February" },
      {"North", "Apple",  7000,  "August"   },
      {"North", "Orange", 7000,  "July"     },
      {"North", "Apple",  6000,  "June"     },
      {"South", "Grape",  8000,  "September"},
      {"West",  "Apple",  3000,  "October"  },
      {"South", "Orange", 10000, "November" },
      {"West",  "Grape",  4000,  "July"     },
      {"North", "Orange", 5000,  "August"   },
      {"East",  "Orange", 1000,  "November" },
      {"East",  "Orange", 4000,  "October"  },
      {"North", "Grape",  5000,  "August"   },
      {"East",  "Apple",  1000,  "December" },
      {"South", "Apple",  10000, "March"    },
      {"East",  "Grape",  7000,  "October"  },
      {"West",  "Grape",  1000,  "September"},
      {"East",  "Grape",  10000, "October"  },
      {"South", "Orange", 8000,  "March"    },
      {"North", "Apple",  4000,  "July"     },
      {"South", "Orange", 5000,  "July"     },
      {"West",  "Apple",  4000,  "June"     },
      {"East",  "Apple",  5000,  "April"    },
      {"North", "Pear",   3000,  "August"   },
      {"East",  "Grape",  9000,  "November" },
      {"North", "Orange", 8000,  "October"  },
      {"East",  "Apple",  10000, "June"     },
      {"South", "Pear",   1000,  "December" },
      {"North", "Grape",  10000, "July"     },
      {"East",  "Grape",  6000,  "February" }
  };

  xwpp::row_col_options_t hidden = {.hidden_ = true};

  xwpp::format_t* header = workbook.add_format();
  header->set_bold();

  // Example 1. Autofilter without conditions.
  // Set up the worksheet data.
  write_worksheet_header(worksheet1, header);

  // Write the row data.
  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet1.write_string(i + 1, 0, region);
    worksheet1.write_string(i + 1, 1, item);
    worksheet1.write_number(i + 1, 2, volume);
    worksheet1.write_string(i + 1, 3, month);
    i++;
  }

  // Add the autofilter.
  worksheet1.autofilter(0, 0, 50, 3);

  // Example 2. Autofilter with a filter condition in the first column.
  // Set up the worksheet data.
  write_worksheet_header(worksheet2, header);

  // Write the row data.
  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet2.write_string(i + 1, 0, region);
    worksheet2.write_string(i + 1, 1, item);
    worksheet2.write_number(i + 1, 2, volume);
    worksheet2.write_string(i + 1, 3, month);

    /* It isn't sufficient to just apply the filter condition below. We
     * must also hide the rows that don't match the criteria since Excel
     * doesn't do that automatically. */
    if(region == "East")
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // Hide rows that don't match the filter.
      worksheet2.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }

    /* Note, the if() statement above is written to match the logic of the
     * criteria in worksheet::filter_column() below. However you could get
     * the same results with the following simpler, but reversed, code:
     *
     *     if (strcmp(data[i].region, "East") != 0) {
     *         worksheet_set_row_opt(worksheet2, i + 1, LXW_DEF_ROW_HEIGHT, NULL, &hidden);
     *     }
     *
     * The same applies to the Examples 3-6 as well.
     */
    i++;
  }

  // Add the autofilter.
  worksheet2.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  xwpp::filter_rule_t filter_rule2 = {.criteria_ = xwpp::filter_criteria_t::EQUAL_TO, .value_string_ = "East"};
  worksheet2.filter_column(0, filter_rule2);

  // Example 3. Autofilter with a dual filter condition in one of the columns.
  // Set up the worksheet data.
  write_worksheet_header(worksheet3, header);

  // Write the row data.
  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet3.write_string(i + 1, 0, region);
    worksheet3.write_string(i + 1, 1, item);
    worksheet3.write_number(i + 1, 2, volume);
    worksheet3.write_string(i + 1, 3, month);

    if(region == "East" || region == "South")
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet3.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }
    i++;
  }

  // Add the autofilter.
  worksheet3.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  xwpp::filter_rule_t filter_rule3a = {.criteria_ = xwpp::filter_criteria_t::EQUAL_TO, .value_string_ = "East"};

  xwpp::filter_rule_t filter_rule3b = {.criteria_ = xwpp::filter_criteria_t::EQUAL_TO, .value_string_ = "South"};

  worksheet3.filter_column2(0, filter_rule3a, filter_rule3b, xwpp::filter_type_t::OR);

  // Example 4. Autofilter with filter conditions in two columns.
  // Set up the worksheet data.
  write_worksheet_header(worksheet4, header);

  // Write the row data.
  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet4.write_string(i + 1, 0, region);
    worksheet4.write_string(i + 1, 1, item);
    worksheet4.write_number(i + 1, 2, volume);
    worksheet4.write_string(i + 1, 3, month);

    if(region == "East" && volume > 3000 && volume < 8000)
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet4.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }
    i++;
  }

  // Add the autofilter.
  worksheet4.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  xwpp::filter_rule_t filter_rule4a = {.criteria_ = xwpp::filter_criteria_t::EQUAL_TO, .value_string_ = "East"};

  xwpp::filter_rule_t filter_rule4b = {.criteria_ = xwpp::filter_criteria_t::GREATER_THAN, .value_ = 3000};

  xwpp::filter_rule_t filter_rule4c = {.criteria_ = xwpp::filter_criteria_t::LESS_THAN, .value_ = 8000};

  worksheet4.filter_column(0, filter_rule4a);
  worksheet4.filter_column2(2, filter_rule4b, filter_rule4c, xwpp::filter_type_t::AND);

  // Example 5. Autofilter with a dual filter condition in one of the columns.
  // Set up the worksheet data.
  write_worksheet_header(worksheet5, header);

  // Write the row data.
  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet5.write_string(i + 1, 0, region);
    worksheet5.write_string(i + 1, 1, item);
    worksheet5.write_number(i + 1, 2, volume);
    worksheet5.write_string(i + 1, 3, month);

    if(region == "East" || region == "North" || region == "South")
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet5.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }
    i++;
  }

  // Add the autofilter.
  worksheet5.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  const std::vector<std::string> list = {"East", "North", "South"};
  worksheet5.filter_list(0, list);

  // Example 6. Autofilter with filter for blanks.
  // Set up the worksheet data.
  write_worksheet_header(worksheet6, header);

  // Simulate one blank cell in the data, to test the filter.
  data[5].region.clear();

  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet6.write_string(i + 1, 0, region);
    worksheet6.write_string(i + 1, 1, item);
    worksheet6.write_number(i + 1, 2, volume);
    worksheet6.write_string(i + 1, 3, month);

    if(region.empty())
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet6.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }
    i++;
  }

  // Add the autofilter.
  worksheet6.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  xwpp::filter_rule_t filter_rule6 = {.criteria_ = xwpp::filter_criteria_t::BLANKS};
  worksheet6.filter_column(0, filter_rule6);

  // Example 7. Autofilter with filter for non-blanks.
  // Set up the worksheet data.
  write_worksheet_header(worksheet7, header);

  for(uint16_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet7.write_string(i + 1, 0, region);
    worksheet7.write_string(i + 1, 1, item);
    worksheet7.write_number(i + 1, 2, volume);
    worksheet7.write_string(i + 1, 3, month);

    if(!region.empty())
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet7.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }
    i++;
  }

  // Add the autofilter.
  worksheet7.autofilter(0, 0, 50, 3);

  // Add the filter criteria.
  xwpp::filter_rule_t filter_rule7 = {.criteria_ = xwpp::filter_criteria_t::NON_BLANKS};
  worksheet7.filter_column(0, filter_rule7);

  workbook.save("autofilter.xlsx");
}
