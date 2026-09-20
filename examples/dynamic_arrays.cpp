/*
 * An example of how to use Xlsxwriter++ to write functions that create
 * dynamic arrays. These functions are new to Excel 365. The examples mirror
 * the examples in the Excel documentation on these functions.
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

// A simple function and data structure to populate some of the worksheets.
struct worksheet_data
{
  std::string col1_;
  std::string col2_;
  std::string col3_;
  int col4_{0};
};

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* header)
{
  const std::vector<worksheet_data> data{
    {.col1_ = "East",  .col2_ = "Tom",    .col3_ = "Apple",  .col4_ = 6380},
    {.col1_ = "West",  .col2_ = "Fred",   .col3_ = "Grape",  .col4_ = 5619},
    {.col1_ = "North", .col2_ = "Amy",    .col3_ = "Pear",   .col4_ = 4565},
    {.col1_ = "South", .col2_ = "Sal",    .col3_ = "Banana", .col4_ = 5323},
    {.col1_ = "East",  .col2_ = "Fritz",  .col3_ = "Apple",  .col4_ = 4394},
    {.col1_ = "West",  .col2_ = "Sravan", .col3_ = "Grape",  .col4_ = 7195},
    {.col1_ = "North", .col2_ = "Xi",     .col3_ = "Pear",   .col4_ = 5231},
    {.col1_ = "South", .col2_ = "Hector", .col3_ = "Banana", .col4_ = 2427},
    {.col1_ = "East",  .col2_ = "Tom",    .col3_ = "Banana", .col4_ = 4213},
    {.col1_ = "West",  .col2_ = "Fred",   .col3_ = "Pear",   .col4_ = 3239},
    {.col1_ = "North", .col2_ = "Amy",    .col3_ = "Grape",  .col4_ = 6520},
    {.col1_ = "South", .col2_ = "Sal",    .col3_ = "Apple",  .col4_ = 1310},
    {.col1_ = "East",  .col2_ = "Fritz",  .col3_ = "Banana", .col4_ = 6274},
    {.col1_ = "West",  .col2_ = "Sravan", .col3_ = "Pear",   .col4_ = 4894},
    {.col1_ = "North", .col2_ = "Xi",     .col3_ = "Grape",  .col4_ = 7580},
    {.col1_ = "South", .col2_ = "Hector", .col3_ = "Apple",  .col4_ = 9814},
  };

  worksheet.write("A1", "Region", header);
  worksheet.write("B1", "Sales Rep", header);
  worksheet.write("C1", "Product", header);
  worksheet.write("D1", "Units", header);

  for(xwpp::row_num_t row_num{1}; const auto& value: data)
  {
    worksheet.write(row_num, 0, value.col1_);
    worksheet.write(row_num, 1, value.col2_);
    worksheet.write(row_num, 2, value.col3_);
    worksheet.write(row_num, 3, value.col4_);
    row_num++;
  }
}

}

int main()
{
  xwpp::workbook_t workbook;

  xwpp::format_t* header1 = workbook.add_format();
  header1->set_bg_color(xwpp::color_t{0x74AC4C});
  header1->set_font_color(xwpp::color_t{0xFFFFFF});

  xwpp::format_t* header2 = workbook.add_format();
  header2->set_bg_color(xwpp::color_t{0x528FD3});
  header2->set_font_color(xwpp::color_t{0xFFFFFF});

  // Example of using the FILTER() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("F2", "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)");

    // Write the data the function will work on.
    worksheet.write("K1", "Product", header2);
    worksheet.write("K2", "Apple");
    worksheet.write("F1", "Region", header2);
    worksheet.write("G1", "Sales Rep", header2);
    worksheet.write("H1", "Product", header2);
    worksheet.write("I1", "Units", header2);

    write_worksheet_data(worksheet, header1);
    worksheet.set_column_pixels("E:E", 20);
    worksheet.set_column_pixels("J:J", 20);
  }

  // Example of using the UNIQUE() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("F2", "=_xlfn.UNIQUE(B2:B17)");

    // A more complex example combining SORT and UNIQUE.
    worksheet.write_dynamic_formula("H2", "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))");

    // Write the data the function will work on.
    worksheet.write("F1", "Sales Rep", header2);
    worksheet.write("H1", "Sales Rep", header2);

    write_worksheet_data(worksheet, header1);
    worksheet.set_column_pixels("E:E", 20);
    worksheet.set_column_pixels("G:G", 20);
  }

  // Example of using the SORT() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("F2", "=_xlfn._xlws.SORT(B2:B17)");

    // A more complex example combining SORT and FILTER.
    worksheet.write_dynamic_formula("H2", "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)");

    // Write the data the function will work on.
    worksheet.write("F1", "Sales Rep", header2);
    worksheet.write("H1", "Product", header2);
    worksheet.write("I1", "Units", header2);

    write_worksheet_data(worksheet, header1);
    worksheet.set_column_pixels("E:E", 20);
    worksheet.set_column_pixels("G:G", 20);
  }

  // Example of using the SORTBY() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("D2", "=_xlfn.SORTBY(A2:B9,B2:B9)");

    // Write the data the function will work on.
    worksheet.write("A1", "Name", header1);
    worksheet.write("B1", "Age", header1);

    worksheet.write("A2", "Tom");
    worksheet.write("A3", "Fred");
    worksheet.write("A4", "Amy");
    worksheet.write("A5", "Sal");
    worksheet.write("A6", "Fritz");
    worksheet.write("A7", "Srivan");
    worksheet.write("A8", "Xi");
    worksheet.write("A9", "Hector");

    worksheet.write("B2", 52);
    worksheet.write("B3", 65);
    worksheet.write("B4", 22);
    worksheet.write("B5", 73);
    worksheet.write("B6", 19);
    worksheet.write("B7", 39);
    worksheet.write("B8", 19);
    worksheet.write("B9", 66);

    worksheet.write("D1", "Name", header2);
    worksheet.write("E1", "Age", header2);

    worksheet.set_column_pixels("C:C", 20);
  }

  // Example of using the XLOOKUP() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("F1", "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)");

    // Write the data the function will work on.
    worksheet.write("A1", "Country", header1);
    worksheet.write("B1", "Abr", header1);
    worksheet.write("C1", "Prefix", header1);

    worksheet.write("A2", "China");
    worksheet.write("A3", "India");
    worksheet.write("A4", "United States");
    worksheet.write("A5", "Indonesia");
    worksheet.write("A6", "Brazil");
    worksheet.write("A7", "Pakistan");
    worksheet.write("A8", "Nigeria");
    worksheet.write("A9", "Bangladesh");

    worksheet.write("B2", "CN");
    worksheet.write("B3", "IN");
    worksheet.write("B4", "US");
    worksheet.write("B5", "ID");
    worksheet.write("B6", "BR");
    worksheet.write("B7", "PK");
    worksheet.write("B8", "NG");
    worksheet.write("B9", "BD");

    worksheet.write("C2", 86);
    worksheet.write("C3", 91);
    worksheet.write("C4", 1);
    worksheet.write("C5", 62);
    worksheet.write("C6", 55);
    worksheet.write("C7", 92);
    worksheet.write("C8", 234);
    worksheet.write("C9", 880);

    worksheet.write("E1", "Brazil", header2);

    worksheet.set_column_pixels("A:A", 100);
    worksheet.set_column_pixels("D:D", 20);
  }

  // Example of using the XMATCH() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("D2", "=_xlfn.XMATCH(C2,A2:A6)");

    // Write the data the function will work on.
    worksheet.write("A1", "Product", header1);

    worksheet.write("A2", "Apple");
    worksheet.write("A3", "Grape");
    worksheet.write("A4", "Pear");
    worksheet.write("A5", "Banana");
    worksheet.write("A6", "Cherry");

    worksheet.write("C1", "Product", header2);
    worksheet.write("D1", "Position", header2);
    worksheet.write("C2", "Grape");

    worksheet.set_column_pixels("B:B", 20);
  }

  // Example of using the RANDARRAY() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("A1", "=_xlfn.RANDARRAY(5,3,1,100, TRUE)");
  }

  // Example of using the SEQUENCE() function.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("A1", "=_xlfn.SEQUENCE(4,5)");
  }

  // Example of using the Spill range operator.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_formula("H2", "=_xlfn.ANCHORARRAY(F2)");
    worksheet.write_dynamic_formula("J2", "=COUNTA(_xlfn.ANCHORARRAY(F2))");

    // Write the data the function will work on.
    worksheet.write_dynamic_formula("F2", "=_xlfn.UNIQUE(B2:B17)");

    worksheet.write("F1", "Unique", header2);
    worksheet.write("H1", "Spill", header2);
    worksheet.write("J1", "Spill", header2);

    write_worksheet_data(worksheet, header1);
    worksheet.set_column_pixels("E:E", 20);
    worksheet.set_column_pixels("G:G", 20);
    worksheet.set_column_pixels("I:I", 20);
  }

  // Example of using dynamic ranges with older Excel functions.
  {
    xwpp::worksheet_t& worksheet = workbook.add_worksheet("Filter");

    worksheet.write_dynamic_array_formula("B1:B3", "=LEN(A1:A3)");

    // Write the data the function will work on.
    worksheet.write("A1", "Foo");
    worksheet.write("A2", "Food");
    worksheet.write("A3", "Frood");
  }

  workbook.save("dynamic_arrays.xlsx");
}
