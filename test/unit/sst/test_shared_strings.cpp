/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/shared_strings.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(shared_strings)

BOOST_AUTO_TEST_CASE(assemble_xml_file_1)
{
  const std::string expected =
      // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"7\" uniqueCount=\"3\">"
        "<si>"
          "<t>Neptune</t>"
        "</si>"
        "<si>"
          "<t>Mars</t>"
        "</si>"
        "<si>"
          "<t>Venus</t>"
        "</si>"
      "</sst>";
  // clang-format on

  xwpp::shared_strings_t sst;
  (void)sst.get_index("Neptune", false);
  (void)sst.get_index("Neptune", false);
  (void)sst.get_index("Neptune", false);
  (void)sst.get_index("Mars", false);
  (void)sst.get_index("Mars", false);
  (void)sst.get_index("Venus", false);
  (void)sst.get_index("Venus", false);

  BOOST_CHECK_EQUAL(expected, sst.assemble_xml_file());
}

BOOST_AUTO_TEST_CASE(assemble_xml_file_2)
{
  const std::string expected =
      // clang-format off
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">"
        "<si>"
          "<t>abcdefg</t>"
        "</si>"
        "<si>"
          "<t xml:space=\"preserve\">   abcdefg</t>"
        "</si>"
        "<si>"
          "<t xml:space=\"preserve\">abcdefg   </t>"
        "</si>"
      "</sst>";
  // clang-format on

  xwpp::shared_strings_t sst;
  (void)sst.get_index("abcdefg", false);
  (void)sst.get_index("   abcdefg", false);
  (void)sst.get_index("abcdefg   ", false);

  BOOST_CHECK_EQUAL(expected, sst.assemble_xml_file());
}

BOOST_AUTO_TEST_SUITE_END()
