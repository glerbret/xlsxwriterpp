/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(rowcol_to_range)
{
  BOOST_CHECK_EQUAL("A1:B1", xwpp::rowcol_to_range(0, 0, 0, 1));
  BOOST_CHECK_EQUAL("C1:J1", xwpp::rowcol_to_range(0, 2, 0, 9));
  BOOST_CHECK_EQUAL("A2:A3", xwpp::rowcol_to_range(1, 0, 2, 0));
  BOOST_CHECK_EQUAL("A10:Y2", xwpp::rowcol_to_range(9, 0, 1, 24));
  BOOST_CHECK_EQUAL("Z8:AA10", xwpp::rowcol_to_range(7, 25, 9, 26));
  BOOST_CHECK_EQUAL("IU2:IV2", xwpp::rowcol_to_range(1, 254, 1, 255));
  BOOST_CHECK_EQUAL("IW2:XFD1", xwpp::rowcol_to_range(1, 256, 0, 16383));
  BOOST_CHECK_EQUAL("A1:XFE1048577", xwpp::rowcol_to_range(0, 0, 1048576, 16384));
  BOOST_CHECK_EQUAL("XFD1048576:XFE1048577", xwpp::rowcol_to_range(1048575, 16383, 1048576, 16384));

  // Test ranges that resolve to single cells.
  BOOST_CHECK_EQUAL("A1", xwpp::rowcol_to_range(0, 0, 0, 0));
  BOOST_CHECK_EQUAL("XFE1048577", xwpp::rowcol_to_range(1048576, 16384, 1048576, 16384));
}

BOOST_AUTO_TEST_CASE(rowcol_to_range_abs)
{
  BOOST_CHECK_EQUAL("$A$1:$B$1", xwpp::rowcol_to_range_abs(0, 0, 0, 1));
  BOOST_CHECK_EQUAL("$C$1:$J$1", xwpp::rowcol_to_range_abs(0, 2, 0, 9));
  BOOST_CHECK_EQUAL("$A$2:$A$3", xwpp::rowcol_to_range_abs(1, 0, 2, 0));
  BOOST_CHECK_EQUAL("$A$10:$Y$2", xwpp::rowcol_to_range_abs(9, 0, 1, 24));
  BOOST_CHECK_EQUAL("$Z$8:$AA$10", xwpp::rowcol_to_range_abs(7, 25, 9, 26));
  BOOST_CHECK_EQUAL("$IU$2:$IV$2", xwpp::rowcol_to_range_abs(1, 254, 1, 255));
  BOOST_CHECK_EQUAL("$IW$2:$XFD$1", xwpp::rowcol_to_range_abs(1, 256, 0, 16383));
  BOOST_CHECK_EQUAL("$A$1:$XFE$1048577", xwpp::rowcol_to_range_abs(0, 0, 1048576, 16384));
  BOOST_CHECK_EQUAL("$XFD$1048576:$XFE$1048577", xwpp::rowcol_to_range_abs(1048575, 16383, 1048576, 16384));

  // Test ranges that resolve to single cells.
  BOOST_CHECK_EQUAL("$A$1", xwpp::rowcol_to_range_abs(0, 0, 0, 0));
  BOOST_CHECK_EQUAL("$XFE$1048577", xwpp::rowcol_to_range_abs(1048576, 16384, 1048576, 16384));
}

BOOST_AUTO_TEST_SUITE_END()
