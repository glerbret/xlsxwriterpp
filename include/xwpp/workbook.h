/*
 * Copyright 2026, Grégory Lerbret
 *
 * Xlsxwriter++ is a C++ port of libxlsxwriter (https://libxlsxwriter.github.io/).
 */

/**
 * @page workbook_page The Workbook object
 *
 * @see @ref workbook.h for full details of the functionality.
 *
 * @file
 *
 * @brief Management of Excel XLSX workbook.
 *
 * The Workbook object is the main object exposed by the Xlsxwriter++ library.
 *
 * It represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * @code
 *  #include "xlsxwriterpp.h"
 *
 *  int main()
 *  {
 *    xwpp::workbook_t workbook;
 *    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
 *
 *    worksheet.write_string(0, 0, "Hello Excel");
 *
 *    workbook.save("filename.xlsx");
 *  }
 * @endcode
 *
 * @image html workbook01.png
 *
 */

#ifndef XWPP_WORKBOOK_H
#define XWPP_WORKBOOK_H

#include "xwpp/chart.h"
#include "xwpp/chartsheet.h"
#include "xwpp/common.h"
#include "xwpp/worksheet.h"

#include <chrono>
#include <cstdint>
#include <list>
#include <string_view>
#include <vector>

/// @cond
namespace xwpp
{
/// @endcond

struct defined_name_t
{
  size_t index_{0};
  bool hidden_{false};
  std::string name_;
  std::string app_name_;
  std::string formula_;
  std::string normalised_name_;
  std::string normalised_sheetname_;
};

/**
 * @brief Workbook document properties.
 *
 * `%doc_properties_t` contains the properties of the Excel document.
 *
 * @todo Add "dc:language" and "cp:revision".
 */
struct doc_properties_t
{
  /**
   * @brief The title of the Excel Document.
   *
   * The title of the Excel Document.
   */
  std::string title_;

  /**
   * @brief The subject of the Excel Document.
   *
   * The subject of the Excel Document.
   */
  std::string subject_;

  /**
   * @brief The author of the Excel Document.
   *
   * The author of the Excel Document.
   */
  std::string author_;

  /**
   * @brief The author of modification of the Excel Document.
   *
   * The author of modification of the Excel Document.
   */
  std::string modif_author_;

  /**
   * @brief The manager field of the Excel Document.
   *
   * The manager field of the Excel Document.
   */
  std::string manager_;

  /**
   * @brief The company field of the Excel Document.
   *
   * The company field of the Excel Document.
   */
  std::string company_;

  /**
   * @brief The category of the Excel Document.
   *
   * The category of the Excel Document.
   */
  std::string category_;

  /**
   * @brief The keywords of the Excel Document.
   *
   * The keywords of the Excel Document.
   */
  std::string keywords_;

  /**
   * @brief The comment field of the Excel Document.
   *
   * The comment field of the Excel Document.
   */
  std::string comments_;

  /**
   * @brief The status of the Excel Document.
   *
   * The status of the Excel Document.
   */
  std::string status_;

  /**
   * @brief The hyperlink base URL of the Excel Document.
   *
   * The hyperlink base URL of the Excel Document.
   */
  std::string hyperlink_base_;

  /**
   * @brief The file creation date/time shown in Excel.
   *
   * The file creation date/time shown in Excel.
   *
   * @note This defaults to the current time and date. If you wish to create files that are
   * binary equivalent (for the same input data) then you should set this
   * creation date/time to a known value.
   */
  std::chrono::system_clock::time_point created_;

  /**
   * @brief The file modification date/time shown in Excel.
   *
   * The file modification date/time shown in Excel.
   */
  std::chrono::system_clock::time_point modified_;
};

/**
 * @brief Representation of an Excel workbook.
 *
 * `%workbook_t` represents the entire spreadsheet as you see it in Excel
 * and internally it represents the Excel file as it is written on disk.
 *
 * @todo Add API to configure default format, in particular font.
 */
class workbook_t
{
public:
  /**
   * @brief Create a new workbook.
   *
   * @param use_zip64 Allow ZIP64 extensions when creating the xlsx file zip container.
   *
   * @code
   *  xwpp::workbook_t workbook;
   * @endcode
   */
  explicit workbook_t(bool use_zip64 = false);

  /**
   * @brief Saves the workbook objet in Excel file.
   *
   * @param filename The name of the Excel file to create.
   *
   * The `%save()` function writes the Excel file to disk.
   *
   * @code
   *  workbook.save("filename.xlsx");
   * @endcode
   *
   * @todo Add API with C++ filesystem (std::path)
   */
  void save(std::string_view filename);

  /**
   * @brief Set the document properties such as Title, Author etc.
   *
   * @param properties Document properties to set.
   *
   * The `%set_properties()` function can be used to set the document
   * properties of the Excel file created by `Xlsxwriter++`. These properties
   * are visible when you use the `Office Button -> Prepare -> Properties`
   * option in Excel and are also available to external applications that read
   * or index windows files.
   *
   * The properties that can be set are:
   *
   * - `title`
   * - `subject`
   * - `author`
   * - `manager`
   * - `company`
   * - `category`
   * - `keywords`
   * - `comments`
   * - `hyperlink_base`
   * - `created`
   *
   * The properties are specified via a `doc_properties_t` struct. All the
   * fields are all optional. An example of how to create and pass the
   * properties is:
   *
   * @code
   *  // Create a properties structure and set some of the fields.
   *  doc_properties_t properties{
   *    .title_    = "This is an example spreadsheet",
   *    .subject_  = "With document properties",
   *    .author_   = "John McNamara",
   *    .manager_  = "Dr. Heinz Doofenshmirtz",
   *    .company_  = "of Wolves",
   *    .category_ = "Example spreadsheets",
   *    .keywords_ = "Sample, Example, Properties",
   *    .comments_ = "Created with Xlsxwriter++",
   *    .status_   = "Quo",
   *  };
   *
   *  // Set the properties in the workbook.
   *  workbook.set_properties(properties);
   * @endcode
   *
   * @image html doc_properties.png
   *
   * The `created` parameter sets the file creation date/time shown in
   * Excel. This defaults to the current time and date if not set. If you wish
   * to create files that are binary equivalent (for the same input data) then
   * you should set this creation date/time to a known value.
   *
   */
  void set_properties(const doc_properties_t& properties);

  /**
   * @brief Set a custom document text property.
   *
   * @param name  The name of the custom property.
   * @param value The value of the custom property.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_custom_property()` functions can be used to set one
   * or more custom document text properties not covered by the standard
   * properties in the `set_properties()` function above.
   *
   *  For example:
   *
   * @code
   *  workbook.set_custom_property("Checked by", "Eve");
   * @endcode
   *
   * @image html custom_properties.png
   *
   * There are several `%set_custom_property()` overload functions for each
   * of the custom property types supported by Excel:
   *
   * - text/string
   * - number (integer and float)
   * - datetime
   * - boolean
   *
   * @note The name and value parameters are limited to 255 characters
   * by Excel.
   *
   * @todo Add overload for all integer types (template)
   * @todo Add overload for all float types (template)
   * @todo Add overload with tm, ...
   * @todo Use string_view instead of two overloads (std::string and char*)
   */
  void set_custom_property(std::string_view name, const std::string& value);
  /// @overload
  void set_custom_property(std::string_view name, const char* value);
  /// @overload
  void set_custom_property(std::string_view name, int32_t value);
  /// @overload
  void set_custom_property(std::string_view name, double value);
  /// @overload
  void set_custom_property(std::string_view name, bool value);
  /// @overload
  void set_custom_property(std::string_view name, const std::chrono::system_clock::time_point& value);
  /// @overload
  void set_custom_property(std::string_view name, const datetime_t& value);
  /// @overload
  void set_custom_property(std::string_view name, const std::chrono::year_month_day& value);

  /**
   * @brief Add a recommendation to open the file in "read-only" mode.
   *
   * This function can be used to set the Excel "Read-only Recommended" option
   * that is available when saving a file. This presents the user of the file
   * with an option to open it in "read-only" mode. This means that any changes
   * to the file can't be saved back to the same file and must be saved to a new
   * file. It can be set as follows:
   *
   * @code
   *  workbook.read_only_recommended();
   * @endcode
   *
   * Which will raise a dialog like the following when opening the file:
   *
   * @image html read_only.png
   *
   * @todo Add to ctor.
   */
  void read_only_recommended();

  /**
   * @brief Set the workbook to use the 1904 epoch.
   *
   * The `%use_1904_epoch()` function can be used to set the workbook to
   * use the 1904 epoch instead of the default 1900 epoch.
   *
   * Excel supports two date epochs. The first based on 1900-01-01 is the default
   * for all Windows versions of Excel and for recent versions of Excel for macOS.
   * Older versions of Excel for macOS used a 1904-01-01 epoch. The 1904 epoch can
   * be set for compatibility with older versions of Excel or to work around the
   * Excel limitation of not being able to handle negative times.
   *
   * This function should be called before `add_worksheet()`.
   *
   * @code
   *  workbook.use_1904_epoch();
   * @endcode
   *
   * @todo Move to ctor.
   */
  void use_1904_epoch();

  /**
   * @brief Set the maximal length of URL.
   *
   * @param max_url_length Maximal length of URL.
   *
   * @todo Add to ctor.
   * @todo For test, check if can be removed.
   */
  void set_max_url_length(uint16_t max_url_length);

  /**
   * @brief Set the size of a workbook window.
   *
   * @param width  Width of the window in pixels.
   * @param height Height of the window in pixels.
   *
   * Set the size of a workbook window. This is generally only useful on macOS
   * since Microsoft Windows uses the window size from the last time an Excel file
   * was opened/saved. The default size is 1073 x 644 pixels.
   *
   * The resulting pixel sizes may not exactly match the target screen and
   * resolution since it is based on the original Excel for Windows sizes. Some
   * trial and error may be required to get an exact size.
   *
   * @todo Add to ctor.
   */
  void set_size(uint16_t width, uint16_t height);

  /**
   * @brief Add a new worksheet to a workbook.
   *
   * @param sheetname Worksheet name.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * @return A reference on @ref worksheet.h "Worksheet" object.
   *
   * The `%add_worksheet()` function adds a new worksheet to a workbook.
   *
   * At least one worksheet should be added to a new workbook: The @ref
   * worksheet.h "Worksheet" object is used to write data and configure a
   * worksheet in the workbook.
   *
   * @code
   *  worksheet = workbook.add_worksheet("Foglio2");    // Foglio2
   *  worksheet = workbook.add_worksheet("Data");       // Data
   * @endcode
   *
   * @image html workbook02.png
   *
   * The worksheet name must be a valid Excel worksheet name, i.e:
   *
   * - The name cannot be empty.
   * - The name is less than or equal to 31 UTF-8 characters.
   * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
   * - The name doesn't start or end with an apostrophe.
   * - The name isn't already in use. (Case insensitive).
   *
   * If any of these errors are encountered the function will throw.
   * You can check for valid name using the `validate_sheetname()`
   * function.
   *
   * @note You should also avoid using the worksheet name "History" (case
   * insensitive) which is reserved in English language versions of
   * Excel. Non-English versions may have restrictions on the equivalent word.
   */
  worksheet_t& add_worksheet(std::string_view sheetname);

  /**
   * @brief Add a new worksheet to a workbook.
   *
   * @return A reference on @ref worksheet.h "Worksheet" object.
   *
   * The `%add_worksheet()` function adds a new worksheet to a workbook.
   *
   * At least one worksheet should be added to a new workbook: The @ref
   * worksheet.h "Worksheet" object is used to write data and configure a
   * worksheet in the workbook.
   *
   * The default Excel convention is followed to name sheet, i.e. `Sheet1`, `Sheet2`, etc.:
   *
   * @code
   *  worksheet = workbook.add_worksheet();     // Sheet1
   *  worksheet = workbook.add_worksheet();     // Sheet2
   * @endcode
   *
   * @image html workbook02.png
   */
  worksheet_t& add_worksheet();

  /**
   * @brief Add a new chartsheet to a workbook.
   *
   * @param sheetname Chartsheet name.
   *
   * @return A reference on @ref chartsheet.h "Chartsheet" object.
   *
   * The `%add_chartsheet()` function adds a new chartsheet to a
   * workbook. The @ref chartsheet.h "Chartsheet" object is like a worksheet
   * except it displays a chart instead of cell data.
   *
   * @code
   *  chartsheet = workbook_add_chartsheet(workbook, "My Chart"); // My Chart
   * @endcode
   *
   * @image html chartsheet.png
   *
   * The chartsheet name must be a valid Excel worksheet name, i.e.:
   *
   * - The name cannot be empty.
   * - The name is less than or equal to 31 UTF-8 characters.
   * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
   * - The name doesn't start or end with an apostrophe.
   * - The name isn't already in use. (Case insensitive).
   *
   * If any of these errors are encountered the function will throw.
   * You can check for valid name using the `validate_sheetname()`
   * function.
   *
   * @note You should also avoid using the worksheet name "History" (case
   * insensitive) which is reserved in English language versions of
   * Excel. Non-English versions may have restrictions on the equivalent word.
   *
   * At least one worksheet should be added to a new workbook when creating a
   * chartsheet in order to provide data for the chart. The @ref worksheet.h
   * "Worksheet" object is used to write data and configure a worksheet in the
   * workbook.
   */
  chartsheet_t& add_chartsheet(std::string_view sheetname);

  /**
   * @brief Add a new chartsheet to a workbook.
   *
   * @return A reference on @ref chartsheet.h "Chartsheet" object.
   *
   * The `%add_chartsheet()` function adds a new chartsheet to a
   * workbook. The @ref chartsheet.h "Chartsheet" object is like a worksheet
   * except it displays a chart instead of cell data.
   *
   * The default Excel convention is followed to named sheet, i.e. `Chart1`, `Chart2`, etc.:
   *
   * @code
   *  chartsheet = workbook.add_chartsheet();     // Chart1
   *  chartsheet = workbook.add_chartsheet();     // Chart2
   * @endcode
   *
   * @image html chartsheet.png
   *
   * At least one worksheet should be added to a new workbook when creating a
   * chartsheet in order to provide data for the chart. The @ref worksheet.h
   * "Worksheet" object is used to write data and configure a worksheet in the
   * workbook.
   */
  chartsheet_t& add_chartsheet();

  /**
   * @brief Get a worksheet object from its name.
   *
   * @param name Worksheet name.
   *
   * @return A pointer on @ref worksheet.h "Worksheet" object.
   *
   * This function returns a @ref worksheet.h "Worksheet" object reference based on its name:
   *
   * @code
   *  const xwpp::worksheet_t* worksheet = workbook.get_worksheet_by_name("Sheet1");
   * @endcode
   *
   * @todo Use reference instead of pointer (like `add_worksheet`).
   * @todo Add non-const overload?
   */
  [[nodiscard]] const worksheet_t* get_worksheet_by_name(std::string_view name) const;

  /**
   * @brief Get a chartsheet object from its name.
   *
   * @param name chartsheet name.
   *
   * @return A pointer on @ref chartsheet.h "Chartsheet" object.
   *
   * This function returns a @ref chartsheet.h "Chartsheet" object reference based on its name:
   *
   * @code
   *  const xwpp::chartsheet_t* chartsheet = workbook.get_chartsheet_by_name("Chart1");
   * @endcode
   *
   * @todo Use reference instead of pointer (like `add_chartsheet`).
   * @todo Add non-const overload?
   */
  [[nodiscard]] const chartsheet_t* get_chartsheet_by_name(std::string_view name) const;

  /**
   * @brief Validate a worksheet or chartsheet name.
   *
   * @param sheetname Sheet name to validate.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function is used to validate a worksheet or chartsheet name according
   * to the rules used by Excel:
   *
   * - The name cannot be empty.
   * - The name is less than or equal to 31 UTF-8 characters.
   * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
   * - The name doesn't start or end with an apostrophe.
   * - The name isn't already in use. (Case insensitive, see the note below).
   *
   * @code
   *  workbook.validate_sheetname("Foglio");
   * @endcode
   *
   * This function is called by `add_worksheet()` and
   * `add_chartsheet()` but it can be explicitly called by the user
   * beforehand to ensure that the sheet name is valid.
   *
   * @note You should also avoid using the worksheet name "History" (case
   * insensitive) which is reserved in English language versions of
   * Excel. Non-English versions may have restrictions on the equivalent word.
   *
   * @note This function does an ASCII lowercase string comparison to determine
   * if the sheet name is already in use. It doesn't take UTF-8 characters into
   * account. Thus it would flag "Café" and "café" as a duplicate (just like
   * Excel) but it wouldn't catch "CAFÉ". If you need a full UTF-8 case
   * insensitive check you should use a third party library to implement it.
   *
   * @todo Return boolean to indicate if name is valid (and remove throw).
   */
  void validate_sheetname(std::string_view sheetname) const;

  /**
   * @brief Create a new @ref format.h "Format" object to formats cells in
   * worksheets.
   *
   * @return A pointer to an instance of format_t.
   *
   * The `%add_format()` function can be used to create new @ref
   * format.h "Format" objects which are used to apply formatting to a cell.
   *
   * @code
   *  // Create the Format.
   *  xwpp::format_t* format = workbook.add_format();
   *
   *  // Set some of the format properties.
   *  format.set_bold();
   *  format.set_font_color(xwpp::color_t::RED);
   *
   *  // Use the format to change the text format in a cell.
   *  worksheet.write_string(0, 0, "Hello", format);
   * @endcode
   *
   * See @ref format.h "the Format object" and @ref working_with_formats
   * sections for more details about Format properties and how to set them.
   *
   * @todo Add builder API to create format.
   */
  [[nodiscard]] format_t* add_format();

  /**
   * @brief Get the default URL format used with `worksheet_t::write_url()`.
   *
   * @return A pointer to an instance of format_t that has hyperlink properties set.
   *
   * This function returns a format_t instance that is used for the default
   * blue underline hyperlink in the `worksheet_t::write_url()` function when a
   * format isn't specified:
   *
   * @code
   *  xwpp::format_t* url_format = workbook.get_default_url_format();
   * @endcode
   *
   * The format is the hyperlink style defined by Excel for the default theme.
   * This format is only ever required when overwriting a string URL with
   * data of a different type. See the example below.
   */
  [[nodiscard]] format_t* get_default_url_format() const;

  /**
   * @brief Unset the default URL format.
   *
   * This function unset the default URL format.
   *
   * @code
   *  workbook.unset_default_url_format();
   * @endcode
   */
  void unset_default_url_format();

  /**
   * @brief Create a new chart to be added to a worksheet.
   *
   * @param chart_type The type of chart to be created. See `chart_type_t`.
   *
   * @return A reference to a chart_t object.
   *
   * The `%add_chart()` function creates a new chart object that can
   * be added to a worksheet.
   *
   * @code
   *  // Create a chart object.
   *  xwpp::chart& chart = workbook.add_chart(LXW_CHART_COLUMN);
   *
   *  // Add data series to the chart.
   *  chart.add_series("", "Sheet1!$A$1:$A$5");
   *  chart.add_series("", "Sheet1!$B$1:$B$5");
   *  chart.add_series("", "Sheet1!$C$1:$C$5");
   *
   *  // Insert the chart into the worksheet
   *  worksheet.insert_chart(CELL("B7"), &chart);
   * @endcode
   *
   * The available chart types are defined in `chart_type_t`. The types of
   * charts that are supported are:
   *
   * | Chart type                                     | Description                           |
   * | :--------------------------------------------- | :-----------------------------------  |
   * | `%chart_type_t::AREA`                          | Area chart                            |
   * | `%chart_type_t::AREA_STACKED`                  | Area chart - stacked                  |
   * | `%chart_type_t::AREA_STACKED_PERCENT`          | Area chart - percentage stacked       |
   * | `%chart_type_t::BAR`                           | Bar chart                             |
   * | `%chart_type_t::BAR_STACKED`                   | Bar chart - stacked                   |
   * | `%chart_type_t::BAR_STACKED_PERCENT`           | Bar chart - percentage stacked        |
   * | `%chart_type_t::COLUMN`                        | Column chart                          |
   * | `%chart_type_t::COLUMN_STACKED`                | Column chart - stacked                |
   * | `%chart_type_t::COLUMN_STACKED_PERCENT`        | Column chart - percentage stacked     |
   * | `%chart_type_t::DOUGHNUT`                      | Doughnut chart                        |
   * | `%chart_type_t::LINE`                          | Line chart                            |
   * | `%chart_type_t::LINE_STACKED`                  | Line chart - stacked                  |
   * | `%chart_type_t::LINE_STACKED_PERCENT`          | Line chart - percentage stacked       |
   * | `%chart_type_t::PIE`                           | Pie chart                             |
   * | `%chart_type_t::SCATTER`                       | Scatter chart                         |
   * | `%chart_type_t::SCATTER_STRAIGHT`              | Scatter chart - straight              |
   * | `%chart_type_t::SCATTER_STRAIGHT_WITH_MARKERS` | Scatter chart - straight with markers |
   * | `%chart_type_t::SCATTER_SMOOTH`                | Scatter chart - smooth                |
   * | `%chart_type_t::SCATTER_SMOOTH_WITH_MARKERS`   | Scatter chart - smooth with markers   |
   * | `%chart_type_t::RADAR`                         | Radar chart                           |
   * | `%chart_type_t::RADAR_WITH_MARKERS`            | Radar chart - with markers            |
   * | `%chart_type_t::RADAR_FILLED`                  | Radar chart - filled                  |
   *
   * See @ref chart.h for details.
   */
  [[nodiscard]] chart_t& add_chart(chart_type_t chart_type);

  /**
   * @brief Add a vbaProject binary to the Excel workbook.
   *
   * @param filename The path/filename of the vbaProject.bin file.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%add_vba_project()` function can be used to add macros or
   * functions to a workbook using a binary VBA project file that has been
   * extracted from an existing Excel xlsm file:
   *
   * @code
   *  workbook.add_vba_project("vbaProject.bin");
   * @endcode
   *
   * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
   * have to be `vbaProject.bin`. Any suitable path/name for an existing VBA bin
   * file will do.
   *
   * Once you add a VBA project had been add to an Xlsxwriter++ workbook you
   * should ensure that the file extension is `.xlsm` to prevent Excel from
   * giving a warning when it opens the file.
   *
   * See also @ref working_with_macros
   *
   * @todo Add API with `std::filesystem`.
   */
  void add_vba_project(const std::string& filename);

  /**
   * @brief Add a vbaProject binary and a vbaProjectSignature binary to the Excel
   * workbook.
   *
   * @param vba_project The path/filename of the vbaProject.bin file.
   * @param signature   The path/filename of the vbaProjectSignature.bin file.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%add_signed_vba_project()` function can be used to add
   * digitally signed macros or functions to a workbook.
   *
   * The function adds a binary VBA project file and a binary VBA project signature
   * file that have been extracted from an existing Excel xlsm file with digitally signed macros:
   *
   * @code
   *  workbook.add_signed_vba_project("vbaProject.bin", "vbaProjectSignature.bin");
   * @endcode
   *
   * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
   * have to be `vbaProject.bin`.
   *
   * Any suitable path/name for an existing VBA bin
   * file will do. The same applies for `vbaProjectSignature.bin`.
   *
   * See also @ref working_with_macros
   *
   * @todo Add API with `std::filesystem`.
   */
  void add_signed_vba_project(const std::string& vba_project, const std::string& signature);

  /**
   * @brief Set the VBA name for the workbook.
   *
   * @param name Name of the workbook used by VBA.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * The `%set_vba_name()` function can be used to set the VBA name for
   * the workbook. This is sometimes required when a vbaProject macro included
   * via `%add_vba_project()` refers to the workbook by a name other
   * than `ThisWorkbook`.
   *
   * @code
   *  workbook.set_vba_name("MyWorkbook");
   * @endcode
   *
   * If an Excel VBA name for the workbook isn't specified then Xlsxwriter++
   * will use `ThisWorkbook`.
   *
   * See also @ref working_with_macros.
   */
  void set_vba_name(std::string_view name);

  /**
   * @brief Create a defined name in the workbook to use as a variable.
   *
   * @param name    The defined name.
   * @param formula The cell or range that the defined name refers to.
   *
   * @throw xwpp::xwpp_exception_t.
   *
   * This function is used to defined a name that can be used to represent a
   * value, a single cell or a range of cells in a workbook: These defined names
   * can then be used in formulas:
   *
   * @code
   *  workbook.define_name("Exchange_rate", "=0.96");
   *  worksheet.write_formula(2, 1, "=Exchange_rate");
   * @endcode
   *
   * @image html defined_name.png
   *
   * As in Excel a name defined like this is "global" to the workbook and can be
   * referred to from any worksheet:
   *
   * @code
   *  // Global workbook name.
   *  workbook.define_name("Sales", "=Sheet1!$G$1:$H$10");
   * @endcode
   *
   * It is also possible to define a local/worksheet name by prefixing it with
   * the sheet name using the syntax `'sheetname!definedname'`:
   *
   * @code
   *  // Local worksheet name.
   *  workbook.define_name("Sheet2!Sales", "=Sheet2!$G$1:$G$10");
   * @endcode
   *
   * If the sheet name contains spaces or special characters you must follow the
   * Excel convention and enclose it in single quotes:
   *
   * @code
   *  workbook.define_name("'New Data'!Sales", "=Sheet2!$G$1:$G$10");
   * @endcode
   *
   * @see [Microsoft Office
   * documentation](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64).
   */
  void define_name(const std::string& name, const std::string& formula);

  /**
   * @brief Set default index for format.
   *
   * Set default index for format.
   *
   * @note Only used by tests.
   *
   * @todo Check to remove it.
   */
  void set_default_xf_indices();

  // TODO Only used by packager, should be private.
  [[nodiscard]] std::string assemble_xml_file();

private:
  // packager_t needs to access to workbook fields to generate the XLSX file.
  friend class packager_t;
  // TODO Remove this friend
  friend class rich_value_t;

  // Functions to retrieve index of format. Provided to worksheets as callback.
  int32_t get_xf_index(format_t* format);
  int32_t get_dxf_index(format_t* format);

  // Function to prepare data before packaging.
  void prepare_fonts();
  void prepare_num_formats();
  void prepare_borders();
  void prepare_vml();
  void store_defined_name(std::string_view name, std::string_view app_name, std::string_view formula, size_t index,
                          bool hidden);
  void prepare_defined_names();
  void store_image_type(image_types_t image_type);
  void prepare_drawings();
  void populate_range_dimensions(series_range_t& range) const;
  void populate_range_data_cache(series_range_t& range) const;
  void populate_range(series_range_t& range) const;
  void add_chart_cache_data() const;
  void prepare_tables();
  void prepare_fills();
  void prepare_workbook();

  // Functions to write XML content.
  [[nodiscard]] std::string write_file_version() const;
  [[nodiscard]] std::string write_file_sharing() const;
  [[nodiscard]] std::string write_workbook_pr() const;
  [[nodiscard]] std::string write_workbook_view() const;
  [[nodiscard]] std::string write_book_views() const;
  [[nodiscard]] std::string write_sheets() const;
  [[nodiscard]] static std::string write_defined_name(const defined_name_t& defined_name);
  [[nodiscard]] std::string write_defined_names() const;

  static const size_t XWPP_SHEETNAME_MAX{31};

  // Use list to not invalidate referenced owned by caller in case of insertion of new items.
  std::list<worksheet_t> worksheets_;
  std::list<chartsheet_t> chartsheets_;
  // As the unicity of sheet name is case insensitive, the name is stored as lower case.
  std::map<std::string, worksheet_t*, std::less<>> worksheet_names_;
  std::map<std::string, chartsheet_t*, std::less<>> chartsheet_names_;
  // Complete list of sheets (worksheets and chartsheets) in creation order
  std::vector<sheet_t*> sheets_;
  std::list<chart_t> charts_;
  std::list<format_t> formats_;
  std::list<defined_name_t> defined_names_;
  std::map<std::string, uint32_t, std::less<>> image_md5_;
  std::map<std::string, uint32_t, std::less<>> embedded_image_md5_;
  std::map<std::string, uint32_t, std::less<>> header_image_md5_;
  std::map<std::string, uint32_t, std::less<>> background_md5_;
  std::vector<chart_t*> ordered_charts_;
  shared_strings_t sst_;
  doc_properties_t properties_;
  std::vector<custom_property_t> custom_properties_;
  size_t first_sheet_{0};
  size_t active_sheet_{0};
  uint16_t num_format_count_{0};
  uint16_t drawing_count_{0};
  uint16_t comment_count_{0};
  uint32_t num_embedded_images_{0};
  uint16_t window_width_{16095};
  uint16_t window_height_{9660};
  uint16_t font_count_{0};
  uint16_t border_count_{0};
  uint16_t fill_count_{0};
  uint16_t max_url_length_{2079};
  uint8_t read_only_{0};
  bool has_png_{false};
  bool has_jpeg_{false};
  bool has_bmp_{false};
  bool has_gif_{false};
  bool has_vml_{false};
  bool has_comments_{false};
  bool has_metadata_{false};
  bool has_embedded_images_{false};
  bool has_dynamic_functions_{false};
  bool has_embedded_image_descriptions_{false};
  // TODO Combine with unordered_set to optimize search and assure uniqueness.
  // TODO And encapsule this combination and related functions in a dedicated types.
  std::vector<format_t*> used_xf_formats_;
  std::vector<format_t*> used_dxf_formats_;
  std::string vba_project_;
  std::string vba_project_signature_;
  std::string vba_codename_;
  bool use_1904_epoch_{false};
  format_t* default_url_format_;
  bool use_zip64_{false};
};

/// @cond
}
/// @endcond

#endif
