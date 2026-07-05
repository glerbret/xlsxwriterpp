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

  worksheet.set_column(COLS("A:A"), 0.08);
  worksheet.set_column(COLS("B:B"), 0.17);
  worksheet.set_column(COLS("C:C"), 0.25);
  worksheet.set_column(COLS("D:D"), 0.33);
  worksheet.set_column(COLS("E:E"), 0.42);
  worksheet.set_column(COLS("F:F"), 0.5);
  worksheet.set_column(COLS("G:G"), 0.58);
  worksheet.set_column(COLS("H:H"), 0.67);
  worksheet.set_column(COLS("I:I"), 0.75);
  worksheet.set_column(COLS("J:J"), 0.83);
  worksheet.set_column(COLS("K:K"), 0.92);
  worksheet.set_column(COLS("L:L"), 1);
  worksheet.set_column(COLS("M:M"), 1.14);
  worksheet.set_column(COLS("N:N"), 1.29);
  worksheet.set_column(COLS("O:O"), 1.43);
  worksheet.set_column(COLS("P:P"), 1.57);
  worksheet.set_column(COLS("Q:Q"), 1.71);
  worksheet.set_column(COLS("R:R"), 1.86);
  worksheet.set_column(COLS("S:S"), 2);
  worksheet.set_column(COLS("T:T"), 2.14);
  worksheet.set_column(COLS("U:U"), 2.29);
  worksheet.set_column(COLS("V:V"), 2.43);
  worksheet.set_column(COLS("W:W"), 2.57);
  worksheet.set_column(COLS("X:X"), 2.71);
  worksheet.set_column(COLS("Y:Y"), 2.86);
  worksheet.set_column(COLS("Z:Z"), 3);
  worksheet.set_column(COLS("AB:AB"), 8.57);
  worksheet.set_column(COLS("AC:AC"), 8.71);
  worksheet.set_column(COLS("AD:AD"), 8.86);
  worksheet.set_column(COLS("AE:AE"), 9);
  worksheet.set_column(COLS("AF:AF"), 9.14);
  worksheet.set_column(COLS("AG:AG"), 9.29);

  workbook.save("test_set_column01.xlsx");
}
