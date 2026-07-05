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

  worksheet.set_column(COLS("A:A"), 0.083333333333333);
  worksheet.set_column(COLS("B:B"), 0.166666666666667);
  worksheet.set_column(COLS("C:C"), 0.250000000000000);
  worksheet.set_column(COLS("D:D"), 0.333333333333333);
  worksheet.set_column(COLS("E:E"), 0.416666666666667);
  worksheet.set_column(COLS("F:F"), 0.500000000000000);
  worksheet.set_column(COLS("G:G"), 0.583333333333333);
  worksheet.set_column(COLS("H:H"), 0.666666666666666);
  worksheet.set_column(COLS("I:I"), 0.750000000000000);
  worksheet.set_column(COLS("J:J"), 0.833333333333333);
  worksheet.set_column(COLS("K:K"), 0.916666666666666);
  worksheet.set_column(COLS("L:L"), 1.000000000000000);
  worksheet.set_column(COLS("M:M"), 1.142857142857140);
  worksheet.set_column(COLS("N:N"), 1.285714285714290);
  worksheet.set_column(COLS("O:O"), 1.428571428571430);
  worksheet.set_column(COLS("P:P"), 1.571428571428570);
  worksheet.set_column(COLS("Q:Q"), 1.714285714285710);
  worksheet.set_column(COLS("R:R"), 1.857142857142860);
  worksheet.set_column(COLS("S:S"), 2.000000000000000);
  worksheet.set_column(COLS("T:T"), 2.142857142857140);
  worksheet.set_column(COLS("U:U"), 2.285714285714290);
  worksheet.set_column(COLS("V:V"), 2.428571428571430);
  worksheet.set_column(COLS("W:W"), 2.571428571428570);
  worksheet.set_column(COLS("X:X"), 2.714285714285710);
  worksheet.set_column(COLS("Y:Y"), 2.857142857142860);
  worksheet.set_column(COLS("Z:Z"), 3.000000000000000);
  worksheet.set_column(COLS("AB:AB"), 8.5714285714285);
  worksheet.set_column(COLS("AC:AC"), 8.711428571428570);
  worksheet.set_column(COLS("AD:AD"), 8.857142857142860);
  worksheet.set_column(COLS("AE:AE"), 9.000000000000000);
  worksheet.set_column(COLS("AF:AF"), 9.142857142857140);
  worksheet.set_column(COLS("AG:AG"), 9.285714285714290);

  workbook.save("test_set_column02.xlsx");
}
