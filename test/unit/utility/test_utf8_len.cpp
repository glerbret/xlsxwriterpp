/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

using namespace std::literals;

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(utf8_len)
{
  BOOST_CHECK_EQUAL(5, xwpp::utf8_len("abcde"));
  BOOST_CHECK_EQUAL(5, xwpp::utf8_len("abcde"s));
  BOOST_CHECK_EQUAL(4, xwpp::utf8_len("&CŽŽ"));
  BOOST_CHECK_EQUAL(4, xwpp::utf8_len("&CŽŽ"s));
}

BOOST_AUTO_TEST_SUITE_END()
