/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <string>

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(array_formula01)
{
  const std::string expected =
      // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:C7\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:3\">"
          "<c r=\"A1\">"
            "<f t=\"array\" ref=\"A1\">SUM(B1:C1*B2:C2)</f>"
            "<v>9500</v>"
          "</c>"
          "<c r=\"B1\">"
            "<v>500</v>"
          "</c>"
          "<c r=\"C1\">"
            "<v>300</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:3\">"
          "<c r=\"A2\">"
            "<f t=\"array\" ref=\"A2\">SUM(B1:C1*B2:C2)</f>"
            "<v>9500</v>"
          "</c>"
          "<c r=\"B2\">"
            "<v>10</v>"
          "</c>"
          "<c r=\"C2\">"
            "<v>15</v>"
          "</c>"
        "</row>"
        "<row r=\"5\" spans=\"1:3\">"
          "<c r=\"A5\">"
            "<f t=\"array\" ref=\"A5:A7\">TREND(C5:C7,B5:B7)</f>"
            "<v>22196</v>"
          "</c>"
          "<c r=\"B5\">"
            "<v>1</v>"
          "</c>"
          "<c r=\"C5\">"
            "<v>20234</v>"
          "</c>"
        "</row>"
        "<row r=\"6\" spans=\"1:3\">"
          "<c r=\"A6\">"
            "<v>0</v>"
          "</c>"
          "<c r=\"B6\">"
            "<v>2</v>"
          "</c>"
          "<c r=\"C6\">"
            "<v>21003</v>"
          "</c>"
        "</row>"
        "<row r=\"7\" spans=\"1:3\">"
          "<c r=\"A7\">"
            "<v>0</v>"
          "</c>"
          "<c r=\"B7\">"
            "<v>3</v>"
          "</c>"
          "<c r=\"C7\">"
            "<v>10000</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  worksheet.write_array_formula_num(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", nullptr, 9500);
  worksheet.write_array_formula_num(1, 0, 1, 0, "{=SUM(B1:C1*B2:C2)}", nullptr, 9500);
  worksheet.write_array_formula_num(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", nullptr, 22196);

  worksheet.write_number(0, 1, 500);
  worksheet.write_number(1, 1, 10);
  worksheet.write_number(4, 1, 1);
  worksheet.write_number(5, 1, 2);
  worksheet.write_number(6, 1, 3);

  worksheet.write_number(0, 2, 300);
  worksheet.write_number(1, 2, 15);
  worksheet.write_number(4, 2, 20234);
  worksheet.write_number(5, 2, 21003);
  worksheet.write_number(6, 2, 10000);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
