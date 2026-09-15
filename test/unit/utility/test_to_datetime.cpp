/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

#define BOOST_TEST_DYN_LINK

#include "xwpp/utility.h"

#include <boost/test/unit_test.hpp>

#include <chrono>
#include <ctime>
#include <string>

using namespace std::literals::chrono_literals;

BOOST_AUTO_TEST_SUITE(utility)

BOOST_AUTO_TEST_CASE(to_datetime_chrono)
{
  {
    const xwpp::datetime_t expected{.year_ = 2010, .month_ = 2, .day_ = 25, .hour_ = 12, .min_ = 35, .sec_ = 56.};
    const std::chrono::system_clock::time_point date{std::chrono::sys_days{2010y / std::chrono::February / 25d} + 12h +
                                                     35min + 56s};

    BOOST_CHECK(expected == xwpp::to_datetime(date));
  }

  {
    const xwpp::datetime_t expected{.year_ = 0, .month_ = 0, .day_ = 0, .hour_ = 12, .min_ = 35, .sec_ = 56.};
    const std::chrono::system_clock::time_point date{std::chrono::sys_days{1970y / std::chrono::January / 1d} + 12h +
                                                     35min + 56s};

    BOOST_CHECK(expected == xwpp::to_datetime(date));
  }
}

BOOST_AUTO_TEST_CASE(to_datetime_tm)
{
  const xwpp::datetime_t expected{.year_ = 2010, .month_ = 2, .day_ = 25, .hour_ = 12, .min_ = 35, .sec_ = 56.};
  const struct tm date = {.tm_sec = 56, .tm_min = 35, .tm_hour = 12, .tm_mday = 25, .tm_mon = 1, .tm_year = 110};

  BOOST_CHECK(expected == xwpp::to_datetime(date));
}

BOOST_AUTO_TEST_SUITE_END()
