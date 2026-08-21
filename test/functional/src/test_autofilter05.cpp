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
    std::string region_;
    std::string item_;
    int volume_ = 0;
    std::string month_;
  };

  const std::vector<row> data{
    {.region_ = "East",  .item_ = "Apple",  .volume_ = 9000,  .month_ = "July"     },
    {.region_ = "East",  .item_ = "Apple",  .volume_ = 5000,  .month_ = "July"     },
    {.region_ = "South", .item_ = "Orange", .volume_ = 9000,  .month_ = "September"},
    {.region_ = "North", .item_ = "Apple",  .volume_ = 2000,  .month_ = "November" },
    {.region_ = "West",  .item_ = "Apple",  .volume_ = 9000,  .month_ = "November" },
    {.region_ = "",      .item_ = "Pear",   .volume_ = 7000,  .month_ = "October"  },
    {.region_ = "North", .item_ = "Pear",   .volume_ = 9000,  .month_ = "August"   },
    {.region_ = "West",  .item_ = "Orange", .volume_ = 1000,  .month_ = "December" },
    {.region_ = "West",  .item_ = "Grape",  .volume_ = 1000,  .month_ = "November" },
    {.region_ = "South", .item_ = "Pear",   .volume_ = 10000, .month_ = "April"    },
    {.region_ = "West",  .item_ = "Grape",  .volume_ = 6000,  .month_ = "January"  },
    {.region_ = "South", .item_ = "Orange", .volume_ = 3000,  .month_ = "May"      },
    {.region_ = "North", .item_ = "Apple",  .volume_ = 3000,  .month_ = "December" },
    {.region_ = "South", .item_ = "Apple",  .volume_ = 7000,  .month_ = "February" },
    {.region_ = "West",  .item_ = "Grape",  .volume_ = 1000,  .month_ = "December" },
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 8000,  .month_ = "February" },
    {.region_ = "South", .item_ = "Grape",  .volume_ = 10000, .month_ = "June"     },
    {.region_ = "West",  .item_ = "Pear",   .volume_ = 7000,  .month_ = "December" },
    {.region_ = "South", .item_ = "Apple",  .volume_ = 2000,  .month_ = "October"  },
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 7000,  .month_ = "December" },
    {.region_ = "North", .item_ = "Grape",  .volume_ = 6000,  .month_ = "April"    },
    {.region_ = "East",  .item_ = "Pear",   .volume_ = 8000,  .month_ = "February" },
    {.region_ = "North", .item_ = "Apple",  .volume_ = 7000,  .month_ = "August"   },
    {.region_ = "North", .item_ = "Orange", .volume_ = 7000,  .month_ = "July"     },
    {.region_ = "North", .item_ = "Apple",  .volume_ = 6000,  .month_ = "June"     },
    {.region_ = "South", .item_ = "Grape",  .volume_ = 8000,  .month_ = "September"},
    {.region_ = "West",  .item_ = "Apple",  .volume_ = 3000,  .month_ = "October"  },
    {.region_ = "South", .item_ = "Orange", .volume_ = 10000, .month_ = "November" },
    {.region_ = "West",  .item_ = "Grape",  .volume_ = 4000,  .month_ = "July"     },
    {.region_ = "North", .item_ = "Orange", .volume_ = 5000,  .month_ = "August"   },
    {.region_ = "East",  .item_ = "Orange", .volume_ = 1000,  .month_ = "November" },
    {.region_ = "East",  .item_ = "Orange", .volume_ = 4000,  .month_ = "October"  },
    {.region_ = "North", .item_ = "Grape",  .volume_ = 5000,  .month_ = "August"   },
    {.region_ = "East",  .item_ = "Apple",  .volume_ = 1000,  .month_ = "December" },
    {.region_ = "South", .item_ = "Apple",  .volume_ = 10000, .month_ = "March"    },
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 7000,  .month_ = "October"  },
    {.region_ = "West",  .item_ = "Grape",  .volume_ = 1000,  .month_ = "September"},
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 10000, .month_ = "October"  },
    {.region_ = "South", .item_ = "Orange", .volume_ = 8000,  .month_ = "March"    },
    {.region_ = "North", .item_ = "Apple",  .volume_ = 4000,  .month_ = "July"     },
    {.region_ = "South", .item_ = "Orange", .volume_ = 5000,  .month_ = "July"     },
    {.region_ = "West",  .item_ = "Apple",  .volume_ = 4000,  .month_ = "June"     },
    {.region_ = "East",  .item_ = "Apple",  .volume_ = 5000,  .month_ = "April"    },
    {.region_ = "North", .item_ = "Pear",   .volume_ = 3000,  .month_ = "August"   },
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 9000,  .month_ = "November" },
    {.region_ = "North", .item_ = "Orange", .volume_ = 8000,  .month_ = "October"  },
    {.region_ = "East",  .item_ = "Apple",  .volume_ = 10000, .month_ = "June"     },
    {.region_ = "South", .item_ = "Pear",   .volume_ = 1000,  .month_ = "December" },
    {.region_ = "North", .item_ = "Grape",  .volume_ = 10000, .month_ = "July"     },
    {.region_ = "East",  .item_ = "Grape",  .volume_ = 6000,  .month_ = "February" }
  };

  // Write the column headers.
  worksheet.write_string(0, 0, "Region");
  worksheet.write_string(0, 1, "Item");
  worksheet.write_string(0, 2, "Volume");
  worksheet.write_string(0, 3, "Month");

  const xwpp::row_col_options_t hidden = {.hidden_ = true};

  // Write the row data.
  for(xwpp::row_num_t i = 0; const auto& [region, item, volume, month]: data)
  {
    worksheet.write_string(i + 1, 0, region);
    worksheet.write_string(i + 1, 1, item);
    worksheet.write_number(i + 1, 2, volume);
    worksheet.write_string(i + 1, 3, month);

    if(region.empty())
    {
      // Row matches the filter, no further action required.
    }
    else
    {
      // We need to hide rows that don't match the filter.
      worksheet.set_row(i + 1, xwpp::DEF_ROW_HEIGHT, nullptr, hidden);
    }

    i++;
  }

  worksheet.autofilter(0, 0, 50, 3);

  const xwpp::filter_rule_t filter_rule1 = {.criteria_ = xwpp::filter_criteria_t::BLANKS};
  worksheet.filter_column(0, filter_rule1);

  workbook.save("test_autofilter05.xlsx");
}
