/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/chart.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(chart)

BOOST_AUTO_TEST_CASE(assemble_xml_file)
{
  const std::string expected =
    // clang-format off
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<c:lang val=\"en-US\"/>"
      "<c:chart>"
        "<c:plotArea>"
          "<c:layout/>"
          "<c:barChart>"
            "<c:barDir val=\"bar\"/>"
            "<c:grouping val=\"clustered\"/>"
            "<c:ser>"
              "<c:idx val=\"0\"/>"
              "<c:order val=\"0\"/>"
              "<c:val>"
                "<c:numRef>"
                  "<c:f>Sheet1!$A$1:$A$5</c:f>"
                  "<c:numCache>"
                    "<c:formatCode>General</c:formatCode>"
                    "<c:ptCount val=\"5\"/>"
                    "<c:pt idx=\"0\">"
                      "<c:v>1</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"1\">"
                      "<c:v>2</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"2\">"
                      "<c:v>3</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"3\">"
                      "<c:v>4</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"4\">"
                      "<c:v>5</c:v>"
                    "</c:pt>"
                  "</c:numCache>"
                "</c:numRef>"
              "</c:val>"
            "</c:ser>"
            "<c:ser>"
              "<c:idx val=\"1\"/>"
              "<c:order val=\"1\"/>"
              "<c:val>"
                "<c:numRef>"
                  "<c:f>Sheet1!$B$1:$B$5</c:f>"
                  "<c:numCache>"
                    "<c:formatCode>General</c:formatCode>"
                    "<c:ptCount val=\"5\"/>"
                    "<c:pt idx=\"0\">"
                      "<c:v>2</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"1\">"
                      "<c:v>4</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"2\">"
                      "<c:v>6</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"3\">"
                      "<c:v>8</c:v>"
                    "</c:pt>"
                    "<c:pt idx=\"4\">"
                      "<c:v>10</c:v>"
                    "</c:pt>"
                  "</c:numCache>"
                "</c:numRef>"
              "</c:val>"
            "</c:ser>"
            "<c:axId val=\"50010001\"/>"
            "<c:axId val=\"50010002\"/>"
          "</c:barChart>"
          "<c:catAx>"
            "<c:axId val=\"50010001\"/>"
            "<c:scaling>"
              "<c:orientation val=\"minMax\"/>"
            "</c:scaling>"
            "<c:axPos val=\"l\"/>"
            "<c:tickLblPos val=\"nextTo\"/>"
            "<c:crossAx val=\"50010002\"/>"
            "<c:crosses val=\"autoZero\"/>"
            "<c:auto val=\"1\"/>"
            "<c:lblAlgn val=\"ctr\"/>"
            "<c:lblOffset val=\"100\"/>"
          "</c:catAx>"
          "<c:valAx>"
            "<c:axId val=\"50010002\"/>"
            "<c:scaling>"
              "<c:orientation val=\"minMax\"/>"
            "</c:scaling>"
            "<c:axPos val=\"b\"/>"
            "<c:majorGridlines/>"
            "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>"
            "<c:tickLblPos val=\"nextTo\"/>"
            "<c:crossAx val=\"50010001\"/>"
            "<c:crosses val=\"autoZero\"/>"
            "<c:crossBetween val=\"between\"/>"
          "</c:valAx>"
        "</c:plotArea>"
        "<c:legend>"
          "<c:legendPos val=\"r\"/>"
          "<c:layout/>"
        "</c:legend>"
        "<c:plotVisOnly val=\"1\"/>"
      "</c:chart>"
      "<c:printSettings>"
        "<c:headerFooter/>"
        "<c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
        "<c:pageSetup/>"
      "</c:printSettings>"
    "</c:chartSpace>";
  // clang-format on

  xwpp::chart_t chart{xwpp::chart_type_t::BAR};
  xwpp::chart_series_t& series1    = chart.add_series("", "Sheet1!$A$1:$A$5");
  series1.values_.ignore_cache_    = true;
  series1.values_.num_data_points_ = 5;
  series1.values_.data_cache_      = {
    {.number_ = 1}, {.number_ = 2}, {.number_ = 3}, {.number_ = 4}, {.number_ = 5},
  };
  xwpp::chart_series_t& series2    = chart.add_series("", "Sheet1!$B$1:$B$5");
  series2.values_.ignore_cache_    = true;
  series2.values_.num_data_points_ = 5;
  series2.values_.data_cache_      = {
    {.number_ = 2}, {.number_ = 4}, {.number_ = 6}, {.number_ = 8}, {.number_ = 10},
  };

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
