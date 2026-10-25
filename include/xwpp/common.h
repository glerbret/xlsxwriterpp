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
 * Common functions and defines for the Xlsxwriter++ library.
 */

#ifndef XWPP_COMMON_H
#define XWPP_COMMON_H

#include <chrono>
#include <cstdint>
#include <string>
#include <variant>

/// @cond
namespace xwpp
{
/// @endcond

/** Default Excel column width in pixels. */
const uint32_t DEF_COL_WIDTH_PIXELS{64};

/** Default Excel column height in pixels. */
const uint32_t DEF_ROW_HEIGHT_PIXELS{20};

// Custom number formats start at 0xA4
const uint16_t FIRST_CUSTOM_FORMAT = 0xA4;

/**
 * Integer data type to represent a row value. Equivalent to `uint32_t`.
 *
 * The maximum row in Excel is 1,048,576.
 */
using row_num_t = uint32_t;

/**
 * Integer data type to represent a column value. Equivalent to `uint16_t`.
 *
 * The maximum column in Excel is 16,384.
 */
using col_num_t = uint16_t;

/**
 * @brief Struct to represent a date and time in Excel.
 *
 * Struct to represent a date and time in Excel. See @ref working_with_dates.
 */
struct datetime_t
{
  friend bool operator==(const datetime_t& lhs, const datetime_t& rhs) = default;
  friend bool operator!=(const datetime_t& lhs, const datetime_t& rhs) = default;

  /** Year     : 1900 - 9999 */
  int year{0};
  /** Month    : 1 - 12 */
  int month{0};
  /** Day      : 1 - 31 */
  int day{0};
  /** Hour     : 0 - 23 */
  int hour{0};
  /** Minute   : 0 - 59 */
  int min{0};
  /** Seconds  : 0 - 59.999 */
  double sec{0.};
};

enum class custom_property_types_t
{
  NONE,
  STRING,
  DOUBLE,
  INTEGER,
  BOOLEAN,
  DATETIME,
};

enum class drawing_orientation_t
{
  LANDSCAPE,
  PORTRAIT,
};

/**
 * @brief Options to control the positioning of objects.
 *
 * Options to control the positioning of sheet objects such as images
 * or charts. See @ref working_with_object_positioning.
 */
enum class object_position_t
{
  /** Default positioning for the object. */
  DEFAULT,

  /** Move and size the worksheet object with the cells. */
  MOVE_AND_SIZE,

  /** Move but don't size the worksheet object with the cells. */
  MOVE_DONT_SIZE,

  /** Don't move or size the worksheet object with the cells. */
  DONT_MOVE_DONT_SIZE,

  /** Same as `MOVE_AND_SIZE` except `Xlsxwriter++` applies hidden
   *  cells after the object is inserted. */
  MOVE_AND_SIZE_AFTER
};

const std::string SCHEMA_MS{"http://schemas.microsoft.com/office/2006/relationships"};
const std::string SCHEMA_ROOT{"http://schemas.openxmlformats.org"};
const std::string SCHEMA_DRAWING{SCHEMA_ROOT + "/drawingml/2006"};
const std::string SCHEMA_OFFICEDOC{SCHEMA_ROOT + "/officeDocument/2006"};
const std::string SCHEMA_PACKAGE{SCHEMA_ROOT + "/package/2006/relationships"};
const std::string SCHEMA_DOCUMENT{SCHEMA_ROOT + "/officeDocument/2006/relationships"};
const std::string SCHEMA_CONTENT{SCHEMA_ROOT + "/package/2006/content-types"};

// Define custom property used in workbook.c and custom.c.
struct custom_property_t
{
  custom_property_types_t type{custom_property_types_t::NONE};
  std::string name;
  std::variant<std::string, int32_t, double, bool, datetime_t> value;
};

/// @cond
}
/// @endcond

#endif
