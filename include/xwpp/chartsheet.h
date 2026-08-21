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
 *
 * @todo Add a common parent class to `%chartsheet_t` and `%worksheet_t`
 * instead of embedded `%worksheet_t` to hold common part.
 * @todo Alternative is to use a third class with no hierarchy relationship as data member.
 */
class chartsheet_t
{
public:
  explicit chartsheet_t(const worksheet_init_data_t& init_data);

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
   * @brief Make a chartsheet the active, i.e., visible chartsheet.
   *
   * The `%chartsheet_activate()` function is used to specify which chartsheet
   * is initially visible in a multi-sheet workbook:
   *
   * @code
   *  xwpp::worksheet_t& worksheet = workbook.add_worksheet();
   *  xwpp::chartsheet_t& chartsheet1 = workbook.add_chartsheet();
   *  xwpp::chartsheet_t& chartsheet2 = workbook.add_chartsheet();
   *  xwpp::chartsheet_t& chartsheet3 = workbook.add_chartsheet();
   *
   *  chartsheet3.activate();
   * @endcode
   *
   * @image html chartsheet_activate.png
   *
   * More than one chartsheet can be selected via the `select()`
   * function, see below, however only one chartsheet can be active.
   *
   * The default active chartsheet is the first chartsheet.
   *
   * @see `worksheet_t::activate()`.
   */
  void activate();

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

  /**
   * @brief Set the paper type for printing.
   *
   * @param paper_size The Excel paper format type.
   *
   * This function is used to set the paper format for the printed output of a
   * chartsheet:
   *
   * @code
   *  chartsheet1.set_paper(1);  // US Letter
   *  chartsheet2.set_paper(9);  // A4
   * @endcode
   *
   * If you do not specify a paper type the chartsheet will print using the
   * printer's default paper style.
   *
   * @see `worksheet_t::set_paper()` for a full list of available paper sizes.
   *
   * @todo Use enum for `%paper_size`.
   */
  void set_paper(uint8_t paper_size);

  /**
   * @brief Set the chartsheet margins for the printed page.
   *
   * @param left   Left margin in inches.   Excel default is 0.7.
   * @param right  Right margin in inches.  Excel default is 0.7.
   * @param top    Top margin in inches.    Excel default is 0.75.
   * @param bottom Bottom margin in inches. Excel default is 0.75.
   *
   * The `%set_margins()` function is used to set the margins of the
   * chartsheet when it is printed. The units are in inches. Specifying `-1` for
   * any parameter will give the default Excel value as shown above.
   *
   * @code
   *  chartsheet.set_margins(1.3, 1.2, -1, -1);
   * @endcode
   */
  void set_margins(double left, double right, double top, double bottom);

  /**
   * @brief Set the printed page header caption.
   *
   * @param str     The header string.
   * @param options Header options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * Headers and footers are generated using a string which is a combination of
   * plain text and control characters.
   *
   * @code
   *  chartsheet.set_header("&LHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    | Hello                                                         |
   *  //    |                                                               |
   *
   *
   *  chartsheet.set_header("&CHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                          Hello                                |
   *  //    |                                                               |
   *
   *
   *  chartsheet.set_header("&RHello");
   *
   *  //     ---------------------------------------------------------------
   *  //    |                                                               |
   *  //    |                                                         Hello |
   *  //    |                                                               |
   * @endcode
   *
   * See `worksheet_t::set_header()` for a full explanation of the syntax of
   * Excel's header formatting and control characters.
   *
   * The optional parameter options allow to specify options for the header.
   * Currently, the only available option is the header margin:
   *
   * @code
   *  xwpp::header_footer_options_t header_options = { .margin_ = 0.2 };
   *
   *  chartsheet.set_header("Some text", header_options);
   * @endcode
   */
  void set_header(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_header(const std::string& str);

  /**
   * @brief Set the printed page footer caption.
   *
   * @param str     The footer string.
   * @param options Footer options.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The syntax of this function is the same as `set_header()`.
   */
  void set_footer(const std::string& str, const std::optional<header_footer_options_t>& options);
  /// @overload
  void set_footer(const std::string& str);

  // TODO Add documentation (for test only)
  void set_dpi(uint16_t horizontal_dpi, uint16_t vertical_dpi);

  /**
   * @brief Hide the current chartsheet.
   *
   * The `%hide()` function is used to hide a chartsheet:
   *
   * @code
   *  chartsheet.hide();
   * @endcode
   *
   * You may wish to hide a chartsheet in order to avoid confusing a user with
   * intermediate data or calculations.
   *
   * @image html hide_sheet.png
   *
   * A hidden chartsheet can not be activated or selected so this function is
   * mutually exclusive with the `activate()` and `select()` functions.
   * In addition, since the first chartsheet will default to being
   * the active chartsheet, you cannot hide the first
   * chartsheet without activating another sheet:
   *
   * @code
   *  chartsheet2.activate();
   *  chartsheet1.hide();
   * @endcode
   *
   * @see `worksheet_t::hide()`.
   */
  void hide();

  /**
   * @brief Set the chartsheet zoom factor.
   *
   * @param scale Chartsheet zoom factor.
   *
   * Set the chartsheet zoom factor in the range `10 <= zoom <= 400`:
   *
   * @code
   *  chartsheet.set_zoom(75);
   * @endcode
   *
   * The default zoom factor is 100. It isn't possible to set the zoom to
   * "Selection" because it is calculated by Excel at run-time.
   *
   * @see `worksheet_t::set_zoom()`.
   */
  void set_zoom(uint16_t scale);

  /**
   * @brief Set the color of the chartsheet tab.
   *
   * @param color The tab color.
   *
   * The `%set_tab_color()` function is used to change the color of
   * the chartsheet tab:
   *
   * @code
   *  chartsheet1.set_tab_color(xwpp::color_t::RED);
   *  chartsheet2.set_tab_color(xwpp::color_t::GREEN);
   *  chartsheet3.set_tab_color(xwpp::color_t(0xFF9900)); // Orange.
   * @endcode
   *
   * The color should be an RGB integer value, see @ref working_with_colors.
   *
   * @see `worksheet_t::set_tab_color()`.
   */
  void set_tab_color(color_t color);

  /**
   * @brief Set the page orientation as landscape.
   *
   * This function is used to set the orientation of a chartsheet's printed page
   * to landscape. The default chartsheet orientation is landscape, so this
   * function isn't generally required:
   *
   * @code
   *  chartsheet.set_landscape();
   * @endcode
   */
  void set_landscape();

  /**
   * @brief Set the page orientation as portrait.
   *
   * This function is used to set the orientation of a chartsheet's printed page
   * to portrait:
   *
   * @code
   *  chartsheet.set_portrait();
   * @endcode
   */
  void set_portrait();

  /**
   * @brief Set a chartsheet tab as selected.
   *
   * The `%select()` function is used to indicate that a chartsheet
   * is selected in a multi-sheet workbook:
   *
   * @code
   *  chartsheet1.activate();
   *  chartsheet2.select();
   *  chartsheet3.select();
   * @endcode
   *
   * A selected chartsheet has its tab highlighted. Selecting chartsheets is a
   * way of grouping them together so that, for example, several chartsheets
   * could be printed in one go. A chartsheet that has been activated via the
   * `activate()` function will also appear as selected.
   *
   * @see `worksheet_t::select()`.
   */
  void select();

  /**
   * @brief Set current chartsheet as the first visible sheet tab.
   *
   * The `activate()` function determines which chartsheet is
   * initially selected.  However, if there are a large number of chartsheets the
   * selected chartsheet may not appear on the screen. To avoid this you can
   * select the leftmost visible chartsheet tab using
   * `%set_first_sheet()`:
   *
   * @code
   *  chartsheet19.set_first_sheet(); // First visible chartsheet tab.
   *  chartsheet20.activate();        // First visible chartsheet.
   * @endcode
   *
   * This function is not required very often. The default value is the first
   * chartsheet.
   *
   * @see `worksheet_t::set_first_sheet()`.
   */
  void set_first_sheet();

  // TODO Only used by packager, should be private.
  [[nodiscard]] std::string assemble_xml_file();

private:
  friend class packager_t;
  friend class workbook_t;

  // Retrieve function
  [[nodiscard]] const std::string& get_sheet_name() const;
  [[nodiscard]] uint16_t get_sheet_index() const;

  // Functions to generate XML
  [[nodiscard]] static std::string write_chartsheet();
  [[nodiscard]] std::string write_sheet_pr() const;
  [[nodiscard]] std::string write_sheet_views();
  [[nodiscard]] std::string write_sheet_protection(const protection_obj_t& protection) const;
  [[nodiscard]] std::string write_page_margins() const;
  [[nodiscard]] std::string write_page_setup() const;
  [[nodiscard]] std::string write_header_footer() const;
  [[nodiscard]] std::string write_drawings();

  // TODO To replace with common class
  // Use an embedded worksheet instance to write XML records that are shared with worksheet_t
  worksheet_t worksheet_;
  chart_t* chart_ = nullptr;
  protection_obj_t protection_;
  bool is_protected_ = false;
  std::string name_;
  std::string quoted_name_;
  uint16_t index_         = 0;
  bool hidden_            = false;
  uint16_t* active_sheet_ = nullptr;
  uint16_t* first_sheet_  = nullptr;
};

/// @cond
}
/// @endcond

#endif
