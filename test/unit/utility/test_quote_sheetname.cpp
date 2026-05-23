/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(quote_sheetname_no_quote)
{
  BOOST_CHECK_EQUAL("Sheet1", xwpp::quote_sheetname("Sheet1"));
  BOOST_CHECK_EQUAL("Sheet.2", xwpp::quote_sheetname("Sheet.2"));
  BOOST_CHECK_EQUAL("Sheet_3", xwpp::quote_sheetname("Sheet_3"));
}

BOOST_AUTO_TEST_CASE(quote_sheetname_quote)
{
  BOOST_CHECK_EQUAL("'Sheet 4'", xwpp::quote_sheetname("Sheet 4"));
  BOOST_CHECK_EQUAL("'Sheet!5'", xwpp::quote_sheetname("Sheet!5"));
}

BOOST_AUTO_TEST_CASE(quote_sheetname_escape_quote)
{
  BOOST_CHECK_EQUAL("'Sheet''6'", xwpp::quote_sheetname("Sheet'6"));
  BOOST_CHECK_EQUAL("'a''''''''''''''''''''''''''''''''''''''''''''''''''''''''''b'",
                    xwpp::quote_sheetname("a'''''''''''''''''''''''''''''b"));
}

BOOST_AUTO_TEST_SUITE_END()
