/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 *
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/workbook.h"

#include "xwpp/exception.h"

#include <boost/test/unit_test.hpp>

#include <string>

BOOST_AUTO_TEST_SUITE(test_workbook)

BOOST_AUTO_TEST_CASE(test_validate_sheetname_valid)
{
  const std::string sheetname = "123456789_123456789_123456789_1";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_NO_THROW(workbook.validate_sheetname(sheetname));
}

BOOST_AUTO_TEST_CASE(test_validate_sheetname_too_long)
{
  const std::string sheetname = "123456789_123456789_123456789_12";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(test_validate_sheetname_invalid_char)
{
  const std::string sheetname = "Sheet[1]";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

BOOST_AUTO_TEST_CASE(test_validate_sheetname_unbalanced_single_quote_start)
{
  const std::string sheetname = "'Sheet1";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

/* Test a sheet name that ends with an apostrophe. */
BOOST_AUTO_TEST_CASE(test_validate_sheetname_unbalanced_single_quote_end)
{
  const std::string sheetname = "Sheet1'";
  const xwpp::workbook_t workbook;

  BOOST_REQUIRE_THROW(workbook.validate_sheetname(sheetname), xwpp::xwpp_exception_t);
}

#if 0
/* Test a sheet name that already exists. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name04) {

    const char* sheetname = "Sheet1";

    lxw_workbook *workbook = workbook_new(NULL);
    workbook_add_worksheet(workbook, sheetname);

    lxw_error exp = LXW_ERROR_SHEETNAME_ALREADY_USED;
    lxw_error got = workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test a sheet name that starts with an apostrophe. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name05) {

    const char* sheetname = "'Sheet1";

    lxw_workbook *workbook = workbook_new(NULL);
    lxw_error exp = LXW_ERROR_SHEETNAME_START_END_APOSTROPHE;
    lxw_error got = workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test a sheet name that ends with an apostrophe. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name06) {

    const char* sheetname = "Sheet1'";

    lxw_workbook *workbook = workbook_new(NULL);
    lxw_error exp = LXW_ERROR_SHEETNAME_START_END_APOSTROPHE;
    lxw_error got = workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name07) {

    const char* sheetname = "Sheet1";

    lxw_workbook *workbook = workbook_new(NULL);
    workbook_add_worksheet(workbook, sheetname);

    lxw_error exp = LXW_ERROR_SHEETNAME_ALREADY_USED;
    lxw_error got = workbook_validate_sheet_name(workbook, "sheet1");

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name08) {

    const char* sheetname = "Café";

    lxw_workbook *workbook = workbook_new(NULL);
    workbook_add_worksheet(workbook, sheetname);

    lxw_error exp = LXW_ERROR_SHEETNAME_ALREADY_USED;
    lxw_error got = workbook_validate_sheet_name(workbook, "café");

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test a sheet name that already exists, case insensitive. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name09) {

    const char* sheetname = "abcde";

    lxw_workbook *workbook = workbook_new(NULL);
    workbook_add_worksheet(workbook, sheetname);

    lxw_error exp = LXW_ERROR_SHEETNAME_ALREADY_USED;
    lxw_error got = workbook_validate_sheet_name(workbook, "ABCDE");

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test for empty sheet name. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name10) {

    const char* sheetname = "";

    lxw_workbook *workbook = workbook_new(NULL);
    lxw_error exp = LXW_ERROR_PARAMETER_IS_EMPTY;
    lxw_error got = workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}

/* Test for NULL sheet name. */
BOOST_AUTO_TEST_CASE(workbook, validate_worksheet_name11) {

    const char* sheetname = NULL;

    lxw_workbook *workbook = workbook_new(NULL);
    lxw_error exp = LXW_ERROR_NULL_PARAMETER_IGNORED;
    lxw_error got = workbook_validate_sheet_name(workbook, sheetname);

    ASSERT_EQUAL(exp, got);

    lxw_workbook_free(workbook);
}
#endif

BOOST_AUTO_TEST_SUITE_END()
