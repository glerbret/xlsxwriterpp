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

BOOST_AUTO_TEST_SUITE(worksheet)

BOOST_AUTO_TEST_CASE(condtional_format08)
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
        "<cfRule type=\"timePeriod\" priority=\"1\" timePeriod=\"yesterday\">"
          "<formula>FLOOR(A1,1)=TODAY()-1</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"2\" timePeriod=\"today\">"
          "<formula>FLOOR(A1,1)=TODAY()</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"3\" timePeriod=\"tomorrow\">"
          "<formula>FLOOR(A1,1)=TODAY()+1</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"4\" timePeriod=\"last7Days\">"
          "<formula>AND(TODAY()-FLOOR(A1,1)&lt;=6,FLOOR(A1,1)&lt;=TODAY())</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"5\" timePeriod=\"lastWeek\">"
          "<formula>AND(TODAY()-ROUNDDOWN(A1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)&lt;(WEEKDAY(TODAY())+7))</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"6\" timePeriod=\"thisWeek\">"
          "<formula>AND(TODAY()-ROUNDDOWN(A1,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"7\" timePeriod=\"nextWeek\">"
          "<formula>AND(ROUNDDOWN(A1,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"8\" timePeriod=\"lastMonth\">"
          "<formula>AND(MONTH(A1)=MONTH(TODAY())-1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=1,YEAR(A1)=YEAR(TODAY())-1)))</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"9\" timePeriod=\"thisMonth\">"
          "<formula>AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))</formula>"
        "</cfRule>"
        "<cfRule type=\"timePeriod\" priority=\"10\" timePeriod=\"nextMonth\">"
          "<formula>AND(MONTH(A1)=MONTH(TODAY())+1,OR(YEAR(A1)=YEAR(TODAY()),AND(MONTH(A1)=12,YEAR(A1)=YEAR(TODAY())+1)))</formula>"
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
    .type_     = xwpp::conditional_format_types_t::TIME_PERIOD,
    .criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_YESTERDAY,
  };
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_TODAY;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_TOMORROW;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_LAST_7_DAYS;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_LAST_WEEK;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_THIS_WEEK;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_NEXT_WEEK;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_LAST_MONTH;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_THIS_MONTH;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  conditional_format.type_     = xwpp::conditional_format_types_t::TIME_PERIOD;
  conditional_format.criteria_ = xwpp::conditional_criteria_t::TIME_PERIOD_NEXT_MONTH;
  worksheet.conditional_format_range(RANGE("A1:A4"), conditional_format);

  BOOST_CHECK_EQUAL(expected, worksheet.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
