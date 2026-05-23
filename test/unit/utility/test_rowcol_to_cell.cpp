/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(rowcol_to_cell)
{
  BOOST_CHECK_EQUAL("A1", xwpp::rowcol_to_cell(0, 0));
  BOOST_CHECK_EQUAL("B1", xwpp::rowcol_to_cell(0, 1));
  BOOST_CHECK_EQUAL("C1", xwpp::rowcol_to_cell(0, 2));
  BOOST_CHECK_EQUAL("J1", xwpp::rowcol_to_cell(0, 9));
  BOOST_CHECK_EQUAL("A2", xwpp::rowcol_to_cell(1, 0));
  BOOST_CHECK_EQUAL("A3", xwpp::rowcol_to_cell(2, 0));
  BOOST_CHECK_EQUAL("A10", xwpp::rowcol_to_cell(9, 0));
  BOOST_CHECK_EQUAL("Y2", xwpp::rowcol_to_cell(1, 24));
  BOOST_CHECK_EQUAL("Z8", xwpp::rowcol_to_cell(7, 25));
  BOOST_CHECK_EQUAL("AA10", xwpp::rowcol_to_cell(9, 26));
  BOOST_CHECK_EQUAL("IU2", xwpp::rowcol_to_cell(1, 254));
  BOOST_CHECK_EQUAL("IV2", xwpp::rowcol_to_cell(1, 255));
  BOOST_CHECK_EQUAL("IW2", xwpp::rowcol_to_cell(1, 256));
  BOOST_CHECK_EQUAL("XFD1", xwpp::rowcol_to_cell(0, 16383));
  BOOST_CHECK_EQUAL("XFE1048577", xwpp::rowcol_to_cell(1048576, 16384));
}

/// BOOST_AUTO_TEST_CASE(rowcol_to_cell_abs)
/// {
///   BOOST_CHECK_EQUAL("A1", xwpp::rowcol_to_cell_abs(0, 0, false, false));
///     BOOST_CHECK_EQUAL("B1", xwpp::rowcol_to_cell_abs(0, 1, false, false));
///     BOOST_CHECK_EQUAL("C1", xwpp::rowcol_to_cell_abs(0, 2, false, false));
///     BOOST_CHECK_EQUAL("J1", xwpp::rowcol_to_cell_abs(0, 9, false, false));
///     BOOST_CHECK_EQUAL("A2", xwpp::rowcol_to_cell_abs(1, 0, false, false));
///     BOOST_CHECK_EQUAL("A3", xwpp::rowcol_to_cell_abs(2, 0, false, false));
///     BOOST_CHECK_EQUAL("A10", xwpp::rowcol_to_cell_abs(9, 0, false, false));
///     BOOST_CHECK_EQUAL("Y2", xwpp::rowcol_to_cell_abs(1, 24, false, false));
///     BOOST_CHECK_EQUAL("Z8", xwpp::rowcol_to_cell_abs(7, 25, false, false));
///     BOOST_CHECK_EQUAL("AA10", xwpp::rowcol_to_cell_abs(9, 26, false, false));
///     BOOST_CHECK_EQUAL("IU2", xwpp::rowcol_to_cell_abs(1, 254, false, false));
///     BOOST_CHECK_EQUAL("IV2", xwpp::rowcol_to_cell_abs(1, 255, false, false));
///     BOOST_CHECK_EQUAL("IW2", xwpp::rowcol_to_cell_abs(1, 256, false, false));
///     BOOST_CHECK_EQUAL("XFD1", xwpp::rowcol_to_cell_abs(0, 16383, false, false));
///     BOOST_CHECK_EQUAL("XFE1048577", xwpp::rowcol_to_cell_abs(1048576, 16384, false, false));

///     BOOST_CHECK_EQUAL("A$1", xwpp::rowcol_to_cell_abs(0, 0, true, false));
///     BOOST_CHECK_EQUAL("$A1", xwpp::rowcol_to_cell_abs(0, 0, false, true));
///     BOOST_CHECK_EQUAL("$A$1", xwpp::rowcol_to_cell_abs(0, 0, true, true));

///     BOOST_CHECK_EQUAL("XFE$1048577", xwpp::rowcol_to_cell_abs(1048576, 16384, true, false));
///     BOOST_CHECK_EQUAL("$XFE1048577", xwpp::rowcol_to_cell_abs(1048576, 16384, false, true));
///     BOOST_CHECK_EQUAL("$XFE$1048577", xwpp::rowcol_to_cell_abs(1048576, 16384, true, true));
/// }

BOOST_AUTO_TEST_SUITE_END()
