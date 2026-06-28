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

// A simple function and data structure to populate some of the worksheets.
struct worksheet_data
{
  char col1[10];
  char col2[10];
  char col3[10];
  int col4;
};

void write_worksheet_data(xwpp::worksheet_t& worksheet, const xwpp::format_t* header)
{
  worksheet_data data[160] = {
      {"East",  "Tom",    "Apple",  6380},
      {"West",  "Fred",   "Grape",  5619},
      {"North", "Amy",    "Pear",   4565},
      {"South", "Sal",    "Banana", 5323},
      {"East",  "Fritz",  "Apple",  4394},
      {"West",  "Sravan", "Grape",  7195},
      {"North", "Xi",     "Pear",   5231},
      {"South", "Hector", "Banana", 2427},
      {"East",  "Tom",    "Banana", 4213},
      {"West",  "Fred",   "Pear",   3239},
      {"North", "Amy",    "Grape",  6520},
      {"South", "Sal",    "Apple",  1310},
      {"East",  "Fritz",  "Banana", 6274},
      {"West",  "Sravan", "Pear",   4894},
      {"North", "Xi",     "Grape",  7580},
      {"South", "Hector", "Apple",  9814},
  };

  worksheet.write_string(CELL("A1"), "Region", header);
  worksheet.write_string(CELL("B1"), "Sales Rep", header);
  worksheet.write_string(CELL("C1"), "Product", header);
  worksheet.write_string(CELL("D1"), "Units", header);

  for(int row = 0; row < 16; row++)
  {
    worksheet.write_string(row + 1, 0, data[row].col1);
    worksheet.write_string(row + 1, 1, data[row].col2);
    worksheet.write_string(row + 1, 2, data[row].col3);
    worksheet.write_number(row + 1, 3, data[row].col4);
  }
}

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet1  = workbook.add_worksheet("Filter");
  xwpp::worksheet_t& worksheet2  = workbook.add_worksheet("Unique");
  xwpp::worksheet_t& worksheet3  = workbook.add_worksheet("Sort");
  xwpp::worksheet_t& worksheet4  = workbook.add_worksheet("Sortby");
  xwpp::worksheet_t& worksheet5  = workbook.add_worksheet("Xlookup");
  xwpp::worksheet_t& worksheet6  = workbook.add_worksheet("Xmatch");
  xwpp::worksheet_t& worksheet7  = workbook.add_worksheet("Randarray");
  xwpp::worksheet_t& worksheet8  = workbook.add_worksheet("Sequence");
  xwpp::worksheet_t& worksheet9  = workbook.add_worksheet("Spill ranges");
  xwpp::worksheet_t& worksheet10 = workbook.add_worksheet("Older functions");

  xwpp::format_t* header1 = workbook.add_format();
  header1->set_bg_color(xwpp::color_t(0x74AC4C));
  header1->set_font_color(xwpp::color_t(0xFFFFFF));

  xwpp::format_t* header2 = workbook.add_format();
  header2->set_bg_color(xwpp::color_t(0x528FD3));
  header2->set_font_color(xwpp::color_t(0xFFFFFF));

  // Example of using the FILTER() function.
  worksheet1.write_dynamic_formula(CELL("F2"), "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)");

  // Write the data the function will work on.
  worksheet1.write_string(CELL("K1"), "Product", header2);
  worksheet1.write_string(CELL("K2"), "Apple");
  worksheet1.write_string(CELL("F1"), "Region", header2);
  worksheet1.write_string(CELL("G1"), "Sales Rep", header2);
  worksheet1.write_string(CELL("H1"), "Product", header2);
  worksheet1.write_string(CELL("I1"), "Units", header2);

  write_worksheet_data(worksheet1, header1);
  worksheet1.set_column_pixels(COLS("E:E"), 20);
  worksheet1.set_column_pixels(COLS("J:J"), 20);

  // Example of using the UNIQUE() function.
  worksheet2.write_dynamic_formula(CELL("F2"), "=_xlfn.UNIQUE(B2:B17)");

  // A more complex example combining SORT and UNIQUE.
  worksheet2.write_dynamic_formula(CELL("H2"), "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))");

  // Write the data the function will work on.
  worksheet2.write_string(CELL("F1"), "Sales Rep", header2);
  worksheet2.write_string(CELL("H1"), "Sales Rep", header2);

  write_worksheet_data(worksheet2, header1);
  worksheet2.set_column_pixels(COLS("E:E"), 20);
  worksheet2.set_column_pixels(COLS("G:G"), 20);

  // Example of using the SORT() function.
  worksheet3.write_dynamic_formula(CELL("F2"), "=_xlfn._xlws.SORT(B2:B17)");

  // A more complex example combining SORT and FILTER.
  worksheet3.write_dynamic_formula(CELL("H2"), "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)");

  // Write the data the function will work on.
  worksheet3.write_string(CELL("F1"), "Sales Rep", header2);
  worksheet3.write_string(CELL("H1"), "Product", header2);
  worksheet3.write_string(CELL("I1"), "Units", header2);

  write_worksheet_data(worksheet3, header1);
  worksheet3.set_column_pixels(COLS("E:E"), 20);
  worksheet3.set_column_pixels(COLS("G:G"), 20);

  // Example of using the SORTBY() function.
  worksheet4.write_dynamic_formula(CELL("D2"), "=_xlfn.SORTBY(A2:B9,B2:B9)");

  // Write the data the function will work on.
  worksheet4.write_string(CELL("A1"), "Name", header1);
  worksheet4.write_string(CELL("B1"), "Age", header1);

  worksheet4.write_string(CELL("A2"), "Tom");
  worksheet4.write_string(CELL("A3"), "Fred");
  worksheet4.write_string(CELL("A4"), "Amy");
  worksheet4.write_string(CELL("A5"), "Sal");
  worksheet4.write_string(CELL("A6"), "Fritz");
  worksheet4.write_string(CELL("A7"), "Srivan");
  worksheet4.write_string(CELL("A8"), "Xi");
  worksheet4.write_string(CELL("A9"), "Hector");

  worksheet4.write_number(CELL("B2"), 52);
  worksheet4.write_number(CELL("B3"), 65);
  worksheet4.write_number(CELL("B4"), 22);
  worksheet4.write_number(CELL("B5"), 73);
  worksheet4.write_number(CELL("B6"), 19);
  worksheet4.write_number(CELL("B7"), 39);
  worksheet4.write_number(CELL("B8"), 19);
  worksheet4.write_number(CELL("B9"), 66);

  worksheet4.write_string(CELL("D1"), "Name", header2);
  worksheet4.write_string(CELL("E1"), "Age", header2);

  worksheet4.set_column_pixels(COLS("C:C"), 20);

  // Example of using the XLOOKUP() function.
  worksheet5.write_dynamic_formula(CELL("F1"), "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)");

  // Write the data the function will work on.
  worksheet5.write_string(CELL("A1"), "Country", header1);
  worksheet5.write_string(CELL("B1"), "Abr", header1);
  worksheet5.write_string(CELL("C1"), "Prefix", header1);

  worksheet5.write_string(CELL("A2"), "China");
  worksheet5.write_string(CELL("A3"), "India");
  worksheet5.write_string(CELL("A4"), "United States");
  worksheet5.write_string(CELL("A5"), "Indonesia");
  worksheet5.write_string(CELL("A6"), "Brazil");
  worksheet5.write_string(CELL("A7"), "Pakistan");
  worksheet5.write_string(CELL("A8"), "Nigeria");
  worksheet5.write_string(CELL("A9"), "Bangladesh");

  worksheet5.write_string(CELL("B2"), "CN");
  worksheet5.write_string(CELL("B3"), "IN");
  worksheet5.write_string(CELL("B4"), "US");
  worksheet5.write_string(CELL("B5"), "ID");
  worksheet5.write_string(CELL("B6"), "BR");
  worksheet5.write_string(CELL("B7"), "PK");
  worksheet5.write_string(CELL("B8"), "NG");
  worksheet5.write_string(CELL("B9"), "BD");

  worksheet5.write_number(CELL("C2"), 86);
  worksheet5.write_number(CELL("C3"), 91);
  worksheet5.write_number(CELL("C4"), 1);
  worksheet5.write_number(CELL("C5"), 62);
  worksheet5.write_number(CELL("C6"), 55);
  worksheet5.write_number(CELL("C7"), 92);
  worksheet5.write_number(CELL("C8"), 234);
  worksheet5.write_number(CELL("C9"), 880);

  worksheet5.write_string(CELL("E1"), "Brazil", header2);

  worksheet5.set_column_pixels(COLS("A:A"), 100);
  worksheet5.set_column_pixels(COLS("D:D"), 20);

  // Example of using the XMATCH() function.
  worksheet6.write_dynamic_formula(CELL("D2"), "=_xlfn.XMATCH(C2,A2:A6)");

  // Write the data the function will work on.
  worksheet6.write_string(CELL("A1"), "Product", header1);

  worksheet6.write_string(CELL("A2"), "Apple");
  worksheet6.write_string(CELL("A3"), "Grape");
  worksheet6.write_string(CELL("A4"), "Pear");
  worksheet6.write_string(CELL("A5"), "Banana");
  worksheet6.write_string(CELL("A6"), "Cherry");

  worksheet6.write_string(CELL("C1"), "Product", header2);
  worksheet6.write_string(CELL("D1"), "Position", header2);
  worksheet6.write_string(CELL("C2"), "Grape");

  worksheet6.set_column_pixels(COLS("B:B"), 20);

  // Example of using the RANDARRAY() function.
  worksheet7.write_dynamic_formula(CELL("A1"), "=_xlfn.RANDARRAY(5,3,1,100, TRUE)");

  // Example of using the SEQUENCE() function.
  worksheet8.write_dynamic_formula(CELL("A1"), "=_xlfn.SEQUENCE(4,5)");

  // Example of using the Spill range operator.
  worksheet9.write_dynamic_formula(CELL("H2"), "=_xlfn.ANCHORARRAY(F2)");
  worksheet9.write_dynamic_formula(CELL("J2"), "=COUNTA(_xlfn.ANCHORARRAY(F2))");

  // Write the data the function will work on.
  worksheet9.write_dynamic_formula(CELL("F2"), "=_xlfn.UNIQUE(B2:B17)");

  worksheet9.write_string(CELL("F1"), "Unique", header2);
  worksheet9.write_string(CELL("H1"), "Spill", header2);
  worksheet9.write_string(CELL("J1"), "Spill", header2);

  write_worksheet_data(worksheet9, header1);
  worksheet9.set_column_pixels(COLS("E:E"), 20);
  worksheet9.set_column_pixels(COLS("G:G"), 20);
  worksheet9.set_column_pixels(COLS("I:I"), 20);

  // Example of using dynamic ranges with older Excel functions.
  worksheet10.write_dynamic_array_formula(RANGE("B1:B3"), "=LEN(A1:A3)");

  // Write the data the function will work on.
  worksheet10.write_string(CELL("A1"), "Foo");
  worksheet10.write_string(CELL("A2"), "Food");
  worksheet10.write_string(CELL("A3"), "Frood");

  workbook.save("dynamic_arrays.xlsx");
}
