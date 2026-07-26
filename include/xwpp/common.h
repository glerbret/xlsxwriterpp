/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @file
 *
 * @brief Common functions and defines for the Xlsxwriter++ library.
 *
 */

#ifndef XWPP_COMMON_H
#define XWPP_COMMON_H

#include <chrono>
#include <cstdint>
#include <string>
#include <variant>

namespace xwpp
{

/** Integer data type to represent a row value. Equivalent to `uint32_t`.
 *
 * The maximum row in Excel is 1,048,576.
 */
using row_num_t = uint32_t;

/** Integer data type to represent a column value. Equivalent to `uint16_t`.
 *
 * The maximum column in Excel is 16,384.
 */
using col_num_t = uint16_t;

// TODO Add again (in complement of std::ptime)
/** @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
// typedef struct lxw_datetime {
/** Year     : 1900 - 9999 */
//     int year;
/** Month    : 1 - 12 */
//     int month;
/** Day      : 1 - 31 */
//     int day;
/** Hour     : 0 - 23 */
//     int hour;
/** Minute   : 0 - 59 */
//     int min;
/** Seconds  : 0 - 59.999 */
//     double sec;
//};

enum class custom_property_types_t
{
  NONE,
  STRING,
  DOUBLE,
  INTEGER,
  BOOLEAN,
  DATETIME
};

// Size of MD5 byte arrays.
const size_t MD5_SIZE = 16;

enum class drawing_orientation_t
{
  LANDSCAPE,
  PORTRAIT,
};

const std::string SCHEMA_MS        = "http://schemas.microsoft.com/office/2006/relationships";
const std::string SCHEMA_ROOT      = "http://schemas.openxmlformats.org";
const std::string SCHEMA_DRAWING   = SCHEMA_ROOT + "/drawingml/2006";
const std::string SCHEMA_OFFICEDOC = SCHEMA_ROOT + "/officeDocument/2006";
const std::string SCHEMA_PACKAGE   = SCHEMA_ROOT + "/package/2006/relationships";
const std::string SCHEMA_DOCUMENT  = SCHEMA_ROOT + "/officeDocument/2006/relationships";
const std::string SCHEMA_CONTENT   = SCHEMA_ROOT + "/package/2006/content-types";

// Define custom property used in workbook.c and custom.c.
struct custom_property_t
{
  custom_property_types_t type_;
  std::string name_;
  std::variant<std::string, int32_t, double, bool, std::chrono::system_clock::time_point> value_;
};

}

#endif
