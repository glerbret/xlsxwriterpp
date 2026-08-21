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

BOOST_AUTO_TEST_CASE(legend_default)
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

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_right)
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
  chart.legend_set_position(xwpp::chart_legend_position_t::RIGHT);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_top)
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
          "<c:legendPos val=\"t\"/>"
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
  chart.legend_set_position(xwpp::chart_legend_position_t::TOP);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_left)
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
          "<c:legendPos val=\"l\"/>"
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
  chart.legend_set_position(xwpp::chart_legend_position_t::LEFT);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_bottom)
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
          "<c:legendPos val=\"b\"/>"
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
  chart.legend_set_position(xwpp::chart_legend_position_t::BOTTOM);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_none)
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
  chart.legend_set_position(xwpp::chart_legend_position_t::NONE);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_overlay_right)
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
          "<c:overlay val=\"1\"/>"
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
  chart.legend_set_position(xwpp::chart_legend_position_t::OVERLAY_RIGHT);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(legend_overlay_left)
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
          "<c:legendPos val=\"l\"/>"
          "<c:layout/>"
          "<c:overlay val=\"1\"/>"
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
  chart.legend_set_position(xwpp::chart_legend_position_t::OVERLAY_LEFT);

  BOOST_CHECK_EQUAL(expected, chart.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
