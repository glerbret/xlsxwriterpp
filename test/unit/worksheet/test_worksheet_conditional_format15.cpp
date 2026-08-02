/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"
#include "xwpp/workbook.h"
#include "xwpp/worksheet.h"

#include <string>

#include <boost/test/unit_test.hpp>

using namespace std::literals::chrono_literals;

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(condtional_format15)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<dimension ref=\"A1:A4\"/>"
      "<sheetViews>"
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"/>"
      "</sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<sheetData>"
        "<row r=\"1\" spans=\"1:1\">"
          "<c r=\"A1\">"
            "<v>10</v>"
          "</c>"
        "</row>"
        "<row r=\"2\" spans=\"1:1\">"
          "<c r=\"A2\">"
            "<v>20</v>"
          "</c>"
        "</row>"
        "<row r=\"3\" spans=\"1:1\">"
          "<c r=\"A3\">"
            "<v>30</v>"
          "</c>"
        "</row>"
        "<row r=\"4\" spans=\"1:1\">"
          "<c r=\"A4\">"
            "<v>40</v>"
          "</c>"
        "</row>"
      "</sheetData>"
      "<conditionalFormatting sqref=\"A1:A4\">"
        "<cfRule type=\"expression\" priority=\"1\">"
          "<formula>$A$1&gt;5</formula>"
        "</cfRule>"
        "<cfRule type=\"expression\" priority=\"2\">"
          "<formula>$A$2&lt;80</formula>"
        "</cfRule>"
        "<cfRule type=\"expression\" priority=\"3\">"
          "<formula>\"1+2\"</formula>"
        "</cfRule>"
        "<cfRule type=\"expression\" priority=\"4\">"
          "<formula>$A$3&gt;$A$4</formula>"
        "</cfRule>"
      "</conditionalFormatting>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
    "</worksheet>";
  // clang-format on

  xwpp::workbook_t workbook;
  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
  worksheet.select();

  worksheet.write_number(CELL("A1"), 10);
  worksheet.write_number(CELL("A2"), 20);
  worksheet.write_number(CELL("A3"), 30);
  worksheet.write_number(CELL("A4"), 40);

  xwpp::conditional_format_t conditional_format{
    .type_         = xwpp::conditional_format_types_t::FORMULA,
    .value_string_ = "=$A$1>5",
  };
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_         = xwpp::conditional_format_types_t::FORMULA;
  conditional_format.value_string_ = "=$A$2<80";
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_         = xwpp::conditional_format_types_t::FORMULA;
  conditional_format.value_string_ = "\"1+2\"";
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_         = xwpp::conditional_format_types_t::FORMULA;
  conditional_format.value_string_ = "=$A$3>$A$4";
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
