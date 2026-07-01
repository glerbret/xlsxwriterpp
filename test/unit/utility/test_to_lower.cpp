/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(to_lower)
{
  BOOST_CHECK_EQUAL("abcde", xwpp::to_lower("abcde"));
  BOOST_CHECK_EQUAL("abcde", xwpp::to_lower("ABCDE"));
  BOOST_CHECK_EQUAL("abcde", xwpp::to_lower("aBcDe"));
  BOOST_CHECK_EQUAL("ab!de", xwpp::to_lower("aB!de"));
  BOOST_CHECK_EQUAL("ab de", xwpp::to_lower("aB de"));
}

BOOST_AUTO_TEST_SUITE_END()
