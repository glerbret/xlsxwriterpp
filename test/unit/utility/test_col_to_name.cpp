/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(col_to_name)
{
  BOOST_CHECK_EQUAL("A", xwpp::col_to_name(0, false));
  BOOST_CHECK_EQUAL("B", xwpp::col_to_name(1, false));
  BOOST_CHECK_EQUAL("C", xwpp::col_to_name(2, false));
  BOOST_CHECK_EQUAL("J", xwpp::col_to_name(9, false));
  BOOST_CHECK_EQUAL("Y", xwpp::col_to_name(24, false));
  BOOST_CHECK_EQUAL("Z", xwpp::col_to_name(25, false));
  BOOST_CHECK_EQUAL("AA", xwpp::col_to_name(26, false));
  BOOST_CHECK_EQUAL("IU", xwpp::col_to_name(254, false));
  BOOST_CHECK_EQUAL("IV", xwpp::col_to_name(255, false));
  BOOST_CHECK_EQUAL("IW", xwpp::col_to_name(256, false));
  BOOST_CHECK_EQUAL("XFD", xwpp::col_to_name(16383, false));
  BOOST_CHECK_EQUAL("XFE", xwpp::col_to_name(16384, false));

  BOOST_CHECK_EQUAL("$A", xwpp::col_to_name(0, true));
  BOOST_CHECK_EQUAL("$B", xwpp::col_to_name(1, true));
  BOOST_CHECK_EQUAL("$C", xwpp::col_to_name(2, true));
  BOOST_CHECK_EQUAL("$J", xwpp::col_to_name(9, true));
  BOOST_CHECK_EQUAL("$Y", xwpp::col_to_name(24, true));
  BOOST_CHECK_EQUAL("$Z", xwpp::col_to_name(25, true));
  BOOST_CHECK_EQUAL("$AA", xwpp::col_to_name(26, true));
  BOOST_CHECK_EQUAL("$IU", xwpp::col_to_name(254, true));
  BOOST_CHECK_EQUAL("$IV", xwpp::col_to_name(255, true));
  BOOST_CHECK_EQUAL("$IW", xwpp::col_to_name(256, true));
  BOOST_CHECK_EQUAL("$XFD", xwpp::col_to_name(16383, true));
  BOOST_CHECK_EQUAL("$XFE", xwpp::col_to_name(16384, true));
}

BOOST_AUTO_TEST_SUITE_END()
