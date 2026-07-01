/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(name_to_col)
{
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col(nullptr));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col(""));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("A"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("A1"));
  BOOST_CHECK_EQUAL(1, xwpp::name_to_col("B1"));
  BOOST_CHECK_EQUAL(2, xwpp::name_to_col("C1"));
  BOOST_CHECK_EQUAL(9, xwpp::name_to_col("J1"));
  BOOST_CHECK_EQUAL(24, xwpp::name_to_col("Y1"));
  BOOST_CHECK_EQUAL(25, xwpp::name_to_col("Z1"));
  BOOST_CHECK_EQUAL(26, xwpp::name_to_col("AA1"));
  BOOST_CHECK_EQUAL(254, xwpp::name_to_col("IU1"));
  BOOST_CHECK_EQUAL(255, xwpp::name_to_col("IV1"));
  BOOST_CHECK_EQUAL(256, xwpp::name_to_col("IW1"));
  BOOST_CHECK_EQUAL(16383, xwpp::name_to_col("XFD1"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_col("XFE1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("$A1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("A$1"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col("$A$1"));
}

BOOST_AUTO_TEST_CASE(name_to_col_2)
{
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col_2(nullptr));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col_2(""));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col_2("AAA"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col_2("AAA:"));
  BOOST_CHECK_EQUAL(0, xwpp::name_to_col_2("AAA:A"));
  BOOST_CHECK_EQUAL(1, xwpp::name_to_col_2("AAA:B"));
  BOOST_CHECK_EQUAL(2, xwpp::name_to_col_2("AAA:C"));
  BOOST_CHECK_EQUAL(9, xwpp::name_to_col_2("AAA:J"));
  BOOST_CHECK_EQUAL(24, xwpp::name_to_col_2("AAA:Y"));
  BOOST_CHECK_EQUAL(25, xwpp::name_to_col_2("AAA:Z"));
  BOOST_CHECK_EQUAL(26, xwpp::name_to_col_2("AAA:AA"));
  BOOST_CHECK_EQUAL(254, xwpp::name_to_col_2("AAA:IU"));
  BOOST_CHECK_EQUAL(255, xwpp::name_to_col_2("AAA:IV"));
  BOOST_CHECK_EQUAL(256, xwpp::name_to_col_2("AAA:IW"));
  BOOST_CHECK_EQUAL(16383, xwpp::name_to_col_2("AAA:XFD"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_col_2("AAA:XFE"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_col_2("AAA1:XFE1"));
  BOOST_CHECK_EQUAL(16384, xwpp::name_to_col_2("$AAA:$XFE"));
}

BOOST_AUTO_TEST_SUITE_END()
