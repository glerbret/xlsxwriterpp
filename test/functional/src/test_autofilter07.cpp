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

  struct row
  {
    std::string region;
    std::string item;
    int volume{0};
    std::string month;
  };

  const std::vector<row> data{
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

  // Write the column headers.
  worksheet.write(2, 3, "Region");
  worksheet.write(2, 4, "Item");
  worksheet.write(2, 5, "Volume");
  worksheet.write(2, 6, "Month");

  const xwpp::row_col_options_t hidden{.hidden = true};

  // Write the row data.
  for(xwpp::row_num_t row_num{3}; const auto& [region, item, volume, month]: data)
  {
    worksheet.write(row_num, 3, region);
    worksheet.write(row_num, 4, item);
    worksheet.write(row_num, 5, volume);
    worksheet.write(row_num, 6, month);

    if(region == "East")
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet.set_row(row_num, hidden);
    }

    row_num++;
  }

  worksheet.autofilter(2, 3, 52, 6);

  const xwpp::filter_rule_t filter_rule1{.criteria = xwpp::filter_criteria_t::EQUAL_TO, .value_string = "East"};
  worksheet.filter_column(3, filter_rule1);

  workbook.save("test_autofilter07.xlsx");
}
