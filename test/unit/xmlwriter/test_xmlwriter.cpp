/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/xmlwriter.h"

#include <boost/test/unit_test.hpp>

#include <string>

using namespace std::string_literals;

BOOST_AUTO_TEST_SUITE(xmlwriter)

BOOST_AUTO_TEST_CASE(xml_declaration)
{
  const std::string expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_declaration());
}

BOOST_AUTO_TEST_CASE(xml_start_tag)
{
  const std::string expected = "<foo>";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_start_tag("foo"));
  BOOST_CHECK_EQUAL(expected, xwpp::xml_start_tag("foo"s));
}

BOOST_AUTO_TEST_CASE(xml_start_tag_with_attributes)
{
  const std::string expected = R"(<foo span="8" baz="7">)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_start_tag("foo"s, {
                                                            {"span", "8"},
                                                            {"baz",  "7"},
  }));
}

BOOST_AUTO_TEST_CASE(xml_start_tag_with_attributes_to_escape)
{
  const std::string expected = R"(<foo span="&amp;&lt;&gt;&quot;">)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_start_tag("foo"s, {
                                                            {"span", "&<>\""},
  }));
}

BOOST_AUTO_TEST_CASE(xml_end_tag)
{
  const std::string expected = "</foo>";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_end_tag("foo"));
  BOOST_CHECK_EQUAL(expected, xwpp::xml_end_tag("foo"s));
}

BOOST_AUTO_TEST_CASE(xml_empty_tag)
{
  const std::string expected = "<foo/>";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_empty_tag("foo"));
  BOOST_CHECK_EQUAL(expected, xwpp::xml_empty_tag("foo"s));
}

BOOST_AUTO_TEST_CASE(xml_empty_tag_with_attributes)
{
  const std::string expected = R"(<foo span="8" baz="7"/>)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_empty_tag("foo"s, {
                                                            {"span", "8"},
                                                            {"baz",  "7"},
  }));
}

BOOST_AUTO_TEST_CASE(xml_empty_tag_with_attributes_to_escape)
{
  const std::string expected = R"(<foo span="&amp;&lt;&gt;&quot;"/>)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_empty_tag("foo"s, {
                                                            {"span", "&<>\""},
  }));
}

BOOST_AUTO_TEST_CASE(xml_data_element)
{
  const std::string expected = "<foo>bar</foo>";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo", "bar"));
  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo"s, "bar"s));
}

BOOST_AUTO_TEST_CASE(xml_data_element_no_data)
{
  const std::string expected = "<foo/>";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo"));
  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo"s));
}

BOOST_AUTO_TEST_CASE(xml_data_element_with_attributes)
{
  const std::string expected = R"(<foo span="8">bar</foo>)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo"s, "bar"s,
                                                     {
                                                       {"span", "8"}
  }));
}

BOOST_AUTO_TEST_CASE(xml_data_element_with_escapes)
{
  const std::string expected = R"(<foo span="8">&amp;&lt;&gt;"</foo>)";

  BOOST_CHECK_EQUAL(expected, xwpp::xml_data_element("foo"s, "&<>\""s,
                                                     {
                                                       {"span", "8"}
  }));
}

BOOST_AUTO_TEST_SUITE_END()
