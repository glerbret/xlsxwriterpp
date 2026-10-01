/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/exception.h"
#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(test_datetime_validate)
{
  BOOST_CHECK_EQUAL(true, xwpp::datetime_validate({2025, 10, 30, 21, 07, 0.0}));
  BOOST_CHECK_EQUAL(true, xwpp::datetime_validate({0, 0, 0, 21, 07, 0.0}));
  BOOST_CHECK_EQUAL(true, xwpp::datetime_validate({1899, 12, 31, 21, 07, 0.0}));

  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1800, 10, 30, 21, 07, 0.0}));
  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1900, 13, 30, 21, 07, 0.0}));
  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1900, 10, 32, 21, 07, 0.0}));
  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1900, 1, 1, 24, 07, 0.0}));
  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1900, 1, 1, 21, 60, 0.0}));
  BOOST_CHECK_EQUAL(false, xwpp::datetime_validate({1900, 1, 1, 21, 07, 60.0}));
}

BOOST_AUTO_TEST_SUITE_END()
