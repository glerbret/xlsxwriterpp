/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @file
 *
 * @brief Exceptions for Xlsxwriter++.
 *
 */

#ifndef XWPP_EXCEPTION_H
#define XWPP_EXCEPTION_H

#include <stdexcept>
#include <string>

namespace xwpp
{

/**
 * @brief Base exception of Xlsxwriter++.
 *
 */
class xwpp_exception_t : public std::runtime_error
{
public:
  explicit xwpp_exception_t(const std::string& what_arg)
    : std::runtime_error(what_arg)
  {
  }

  explicit xwpp_exception_t(const char* what_arg)
    : std::runtime_error(what_arg)
  {
  }
};

/**
 * @brief Out of range.
 *
 */
class xwpp_out_of_range_t : public xwpp_exception_t
{
public:
  explicit xwpp_out_of_range_t(const std::string& what_arg)
    : xwpp_exception_t(what_arg)
  {
  }

  explicit xwpp_out_of_range_t(const char* what_arg)
    : xwpp_exception_t(what_arg)
  {
  }
};

}

#endif
