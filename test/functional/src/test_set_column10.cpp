/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#include "xlsxwriterpp.h"

int main()
{
  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();

  worksheet.set_column_pixels(COLS("A:A"), 1);
  worksheet.set_column_pixels(COLS("B:B"), 2);
  worksheet.set_column_pixels(COLS("C:C"), 3);
  worksheet.set_column_pixels(COLS("D:D"), 4);
  worksheet.set_column_pixels(COLS("E:E"), 5);
  worksheet.set_column_pixels(COLS("F:F"), 6);
  worksheet.set_column_pixels(COLS("G:G"), 7);
  worksheet.set_column_pixels(COLS("H:H"), 8);
  worksheet.set_column_pixels(COLS("I:I"), 9);
  worksheet.set_column_pixels(COLS("J:J"), 10);
  worksheet.set_column_pixels(COLS("K:K"), 11);
  worksheet.set_column_pixels(COLS("L:L"), 12);
  worksheet.set_column_pixels(COLS("M:M"), 13);
  worksheet.set_column_pixels(COLS("N:N"), 14);
  worksheet.set_column_pixels(COLS("O:O"), 15);
  worksheet.set_column_pixels(COLS("P:P"), 16);
  worksheet.set_column_pixels(COLS("Q:Q"), 17);
  worksheet.set_column_pixels(COLS("R:R"), 18);
  worksheet.set_column_pixels(COLS("S:S"), 19);
  worksheet.set_column_pixels(COLS("T:T"), 20);
  worksheet.set_column_pixels(COLS("U:U"), 21);
  worksheet.set_column_pixels(COLS("V:V"), 22);
  worksheet.set_column_pixels(COLS("W:W"), 23);
  worksheet.set_column_pixels(COLS("X:X"), 24);
  worksheet.set_column_pixels(COLS("Y:Y"), 25);
  worksheet.set_column_pixels(COLS("Z:Z"), 26);
  worksheet.set_column_pixels(COLS("AB:AB"), 65);
  worksheet.set_column_pixels(COLS("AC:AC"), 66);
  worksheet.set_column_pixels(COLS("AD:AD"), 67);
  worksheet.set_column_pixels(COLS("AE:AE"), 68);
  worksheet.set_column_pixels(COLS("AF:AF"), 69);
  worksheet.set_column_pixels(COLS("AG:AG"), 70);

  workbook.save("test_set_column10.xlsx");
}
