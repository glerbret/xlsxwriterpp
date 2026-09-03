/*
 * chartsheet - A Xlsxwriter++ library for creating Excel XLSX chartsheet files.
 *
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page chartsheet_page The Chartsheet object
 *
 * The Chartsheet object represents an Excel chartsheet, which is a type of
 * worksheet that only contains a chart. The Chartsheet object handles
 * operations such as adding a chart and setting the page layout.
 *
 * @see @ref chartsheet.h for full details of the functionality.
 *
 * @file chartsheet.h
 *
 * @brief Management of an Excel chartsheet.
 *
 * The Chartsheet object represents an Excel chartsheet. It handles operations
 * such as adding a chart and setting the page layout.
 *
 * A Chartsheet object isn't created directly. Instead a chartsheet is created
 * by calling the `workbook_t::add_chartsheet()` function from a Workbook object.
 * A chartsheet object functions as a worksheet and not as a chart. In order to
 * have it display data a `chart_t` object must be created and added to the
 * chartsheet:
 *
 * @code
 *  #include "xlsxwriterpp.h"
 *
 *  int main()
 *  {
 *    xwpp::workbook_t workbook;
 *    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
 *    xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
 *    xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::BAR);
 *
 *    //... Set up the chart.
 *
 *    // Add the chart to the chartsheet.
 *    return workbook.save("chartsheet.xlsx");
 *  }
 * @endcode
 *
 * @image html chartsheet.png
 *
 * The data for the chartsheet chart must be contained on a separate
 * worksheet. That is why it is always created in conjunction with at least
 * one data worksheet, as shown above.
 */

#ifndef XWPP_CHARTSHEET_H
#define XWPP_CHARTSHEET_H

#include "xwpp/common.h"
#include "xwpp/drawing.h"
#include "xwpp/sheet.h"
#include "xwpp/utility.h"
#include "xwpp/worksheet.h"

#include <cstdint>
#include <optional>

/// @cond
namespace xwpp
{
/// @endcond

/**
 * @brief Representation an Excel chartsheet.
 *
 * `%chartsheet_t` represents an Excel chartsheet.
 */
class chartsheet_t : public sheet_t
{
public:
  explicit chartsheet_t(const sheet_init_data_t& init_data);

  /**
   * @brief Insert a chart object into a chartsheet.
   *
   * @param chart   A `chart_t` object created via `workbook_t::add_chart()`.
   * @param options Optional chart parameters.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_chart()` function can be used to insert a chart into a
   * chartsheet. The chart object must be created first using the
   * `workbook_t::add_chart()` function and configured using the @ref chart.h
   * functions.
   *
   * @code
   *  // Create the chartsheet.
   *  xwpp::chartsheet_t& chartsheet = workbook.add_chartsheet();
   *
   *  // Create a chart object.
   *  xwpp::chart_t& chart = workbook.add_chart(xwpp::chart_type_t::LINE);
   *
   *  // Add a data series to the chart.
   *  chart.add_series("", "=Sheet1!$A$1:$A$6");
   *
   *  // Insert the chart into the chartsheet.
   *  chartsheet.set_chart(&chart);
   * @endcode
   *
   * @image html chartsheet2.png
   *
   * @note A chart may only be inserted once into a chartsheet or a worksheet.
   * If several similar charts are required then each one must be created
   * separately.
   */
  void set_chart(chart_t* chart, const std::optional<chart_options_t>& options);
  /// @overload
  void set_chart(chart_t* chart);

  /**
   * @brief Protect elements of a chartsheet from modification.
   *
   * @param password A chartsheet password.
   * @param options  Chartsheet elements to protect.
   *
   * The `%protect()` function protects chartsheet elements from
   * modification:
   *
   * @code
   *  chartsheet.protect("Some Password", options);
   * @endcode
   *
   * The `password` and `options` are both optional:
   *
   * @code
   *  chartsheet2.protect(my_options);
   *  chartsheet3.protect("password");
   *  chartsheet4.protect("password", my_options);
   * @endcode
   *
   * Passing an empty password is the same as turning on protection without a
   * password. Passing an empty password and empty options had no effect on
   * chartsheets.
   *
   * You can specify which chartsheet elements you wish to protect by passing a
   * `protection_t` pointer in the `options` argument. In Excel chartsheets only
   * have two protection options:
   *
   * - `%no_content_`
   * - `%no_objects_`
   *
   * All parameters are off by default. Individual elements can be protected as
   * follows:
   *
   * @code
   *  xwpp::protection_t options{
   *    .no_content_ = true,
   *    .no_objects_ = true,
   *  };
   *
   *  chartsheet.protect(options);
   * @endcode
   *
   * @see `worksheet_t::protect()`.
   *
   * @note Sheet level passwords in Excel offer **very** weak
   * protection. They don't encrypt your data and are very easy to
   * deactivate. Full workbook encryption is not supported by `Xlsxwriter++`
   * since it requires a completely different file format.
   */
  void protect(const std::string& password, std::optional<protection_t> options);
  /// @overload
  void protect(const std::string& password);
  /// @overload
  void protect(std::optional<protection_t> options);
  /// @overload
  void protect();

  [[nodiscard]] std::string assemble_xml_file();

private:
  friend class packager_t;
  friend class workbook_t;

  // Functions to generate XML
  [[nodiscard]] static std::string write_chartsheet();
  [[nodiscard]] std::string write_sheet_view() override;
  [[nodiscard]] std::string write_page_setup() const override;
  [[nodiscard]] std::string write_sheet_pr() const override;

  // Preparation functions called by workbook and/or packager.
  void set_error_cell(const object_properties_t& object_props, uint32_t ref_id) override;
  void prepare_background([[maybe_unused]] uint32_t image_ref_id,
                          [[maybe_unused]] object_properties_t& object_props) override;
  void prepare_image(uint32_t image_ref_id, uint32_t drawing_id, object_properties_t& object_props) override;
  void prepare_chart(uint32_t chart_ref_id, uint32_t drawing_id, object_properties_t& object_props) override;

  chart_t* chart_{nullptr};
  bool is_protected_{false};
};

/// @cond
}
/// @endcond

#endif
