/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

#include <chrono>

using namespace std::literals::chrono_literals;

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(datetime_to_excel_datetime, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(
    0 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1899y / std::chrono::December / 31d} + 0h + 0min + 0s));
  BOOST_TEST(30188.010650613425 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1982y / std::chrono::August / 25d} + 0h + 15min +
                                              20s + 213ms));
  BOOST_TEST(60376.011670023145 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2065y / std::chrono::April / 19d} + 0h + 16min +
                                              48s + 290ms));
  BOOST_TEST(90565.038488958337 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2147y / std::chrono::December / 15d} + 0h + 55min +
                                              25s + 446ms));
  BOOST_TEST(120753.04359827546 ==
             xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2230y / std::chrono::August / 10d} + 1h + 2min +
                                              46s + 891ms));
}

BOOST_AUTO_TEST_CASE(datetime_date_only, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1899y / std::chrono::December / 31d}));
  BOOST_TEST(1461 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{1903y / std::chrono::December / 31d}));
  BOOST_TEST(36708 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2000y / std::chrono::July / 1d}));
  BOOST_TEST(60376 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2065y / std::chrono::April / 19d}));
  BOOST_TEST(120753 == xwpp::datetime_to_excel_datetime(std::chrono::sys_days{2230y / std::chrono::August / 10d}));
}

BOOST_AUTO_TEST_CASE(datetime_date_only_1904, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1904y / std::chrono::January / 1d}, true));
  BOOST_TEST(1159 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1907y / std::chrono::March / 5d}, true));
  BOOST_TEST(34879 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{1999y / std::chrono::June / 30d}, true));
  BOOST_TEST(35702 == xwpp::datetime_to_excel_date_with_epoch(
                        std::chrono::sys_days{2001y / std::chrono::September / 30d}, true));
  BOOST_TEST(58914 ==
             xwpp::datetime_to_excel_date_with_epoch(std::chrono::sys_days{2065y / std::chrono::April / 19d}, true));
}

BOOST_AUTO_TEST_CASE(datetime_time_only, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{0h + 0min + 0s}));
  BOOST_TEST(0.11472348379629631 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{2h + 45min + 12s + 109ms}));
  BOOST_TEST(0.39631690972222228 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{9h + 30min + 41s + 781ms}));
  BOOST_TEST(0.6874536921296297 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{16h + 29min + 55s + 999ms}));
  BOOST_TEST(0.90703436342592603 ==
             xwpp::datetime_to_excel_datetime(std::chrono::system_clock::time_point{21h + 46min + 7s + 769ms}));
}

BOOST_AUTO_TEST_CASE(unixtime, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::unixtime_to_excel_date(-2209075200LL));        // 1899-12-31 00:00
  BOOST_TEST(0.5 == xwpp::unixtime_to_excel_date(-2209032000LL));      // 1899-12-31 12:00
  BOOST_TEST(1 == xwpp::unixtime_to_excel_date(-2208988800LL));        // 1900-01-01 00:00
  BOOST_TEST(1.5 == xwpp::unixtime_to_excel_date(-2208945600LL));      // 1900-01-01 12:00
  BOOST_TEST(59 == xwpp::unixtime_to_excel_date(-2203977600LL));       // 1900-02-28 00:00
  BOOST_TEST(59.5 == xwpp::unixtime_to_excel_date(-2203934400LL));     // 1900-02-28 12:00
  BOOST_TEST(61 == xwpp::unixtime_to_excel_date(-2203891200LL));       // 1900-03-01 00:00
  BOOST_TEST(61.5 == xwpp::unixtime_to_excel_date(-2203848000LL));     // 1900-03-01 12:00
  BOOST_TEST(25569 == xwpp::unixtime_to_excel_date(0));                // 1970-01-01 00:00
  BOOST_TEST(25569.5 == xwpp::unixtime_to_excel_date(43200));          // 1970-01-01 12:00
  BOOST_TEST(36526 == xwpp::unixtime_to_excel_date(946684800));        // 2000-01-01 00:00
  BOOST_TEST(36526.5 == xwpp::unixtime_to_excel_date(946728000));      // 2000-01-01 12:00
  BOOST_TEST(2958101 == xwpp::unixtime_to_excel_date(253370764800));   // 9999-01-01 00:00
  BOOST_TEST(2958101.5 == xwpp::unixtime_to_excel_date(253370808000)); // 9999-01-01 12:00
}

BOOST_AUTO_TEST_CASE(unixtime_1904, *boost::unit_test::tolerance(0.00001))
{
  BOOST_TEST(0 == xwpp::unixtime_to_excel_date_with_epoch(-2082844800, true));     // 1904-01-01 00:00
  BOOST_TEST(0.5 == xwpp::unixtime_to_excel_date_with_epoch(-2082801600, true));   // 1904-01-01 12:00
  BOOST_TEST(24107 == xwpp::unixtime_to_excel_date_with_epoch(0, true));           // 1970-01-01 00:00
  BOOST_TEST(24107.5 == xwpp::unixtime_to_excel_date_with_epoch(43200, true));     // 1970-01-01 12:00
  BOOST_TEST(35064 == xwpp::unixtime_to_excel_date_with_epoch(946684800, true));   // 2000-01-01 00:00
  BOOST_TEST(35064.5 == xwpp::unixtime_to_excel_date_with_epoch(946728000, true)); // 2000-01-01 12:00
}

BOOST_AUTO_TEST_SUITE_END()
