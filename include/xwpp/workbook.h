/*
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2026, Grégory Lerbret
 */

/**
 * @page workbook_page The workbook_t object
 *
 * The workbook_t is the main object exposed by the Xlsxwriter++ library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * See @ref workbook.h for full details of the functionality.
 *
 * @file
 *
 * @brief Functions related to creating an Excel xlsx workbook.
 *
 * The workbook_t is the main object exposed by the Xlsxwriter++ library. It
 * represents the entire spreadsheet as you see it in Excel and internally it
 * represents the Excel file as it is written on disk.
 *
 * @code
 *     #include "xlsxwriterpp.h"
 *
 *     int main() {
 *
 *         lxw_workbook  *workbook  = workbook_new("filename.xlsx");
 *         lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
 *
 *         worksheet_write_string(worksheet, 0, 0, "Hello Excel", NULL);
 *
 *         return workbook_close(workbook);
 *     }
 * @endcode
 *
 * @image html workbook01.png
 *
 */

#ifndef XWPP_WORKBOOK_H
#define XWPP_WORKBOOK_H

#include "xwpp/chart.h"
#include "xwpp/common.h"
#include "xwpp/worksheet.h"

#include <chrono>
#include <cstdint>
#include <list>
#include <string_view>
#include <variant>
#include <vector>

namespace xwpp
{

/// #define LXW_DEFINED_NAME_LENGTH 128

/* Struct to hold the 2 sheet types. */
/// typedef struct lxw_sheet {
///     uint8_t is_chartsheet;

///     union {
///         lxw_worksheet *worksheet;
///         lxw_chartsheet *chartsheet;
///     } u;

///     STAILQ_ENTRY (lxw_sheet) list_pointers;
/// } lxw_sheet;

/* Struct to represent a worksheet name/pointer pair. */
/// typedef struct lxw_worksheet_name {
///     const char *name;
///     lxw_worksheet *worksheet;

///     RB_ENTRY (lxw_worksheet_name) tree_pointers;
/// } lxw_worksheet_name;

/* Struct to represent a chartsheet name/pointer pair. */
/// typedef struct lxw_chartsheet_name {
///     const char *name;
///     lxw_chartsheet *chartsheet;

///     RB_ENTRY (lxw_chartsheet_name) tree_pointers;
/// } lxw_chartsheet_name;

/* Struct to represent a defined name. */
/// typedef struct lxw_defined_name {
///     int16_t index;
///     uint8_t hidden;
///     char name[LXW_DEFINED_NAME_LENGTH];
///     char app_name[LXW_DEFINED_NAME_LENGTH];
///     char formula[LXW_DEFINED_NAME_LENGTH];
///     char normalised_name[LXW_DEFINED_NAME_LENGTH];
///     char normalised_sheetname[LXW_DEFINED_NAME_LENGTH];
/// } lxw_defined_name;

struct doc_properties_t
{
  // TODO Add "dc:language" and "cp:revision"

  /** The title of the Excel Document. */
  std::string title_;

  /** The subject of the Excel Document. */
  std::string subject_;

  /** The author of the Excel Document. */
  std::string author_;

  /** The author of modification of the Excel Document. */
  std::string modif_author_;

  /** The manager field of the Excel Document. */
  std::string manager_;

  /** The company field of the Excel Document. */
  std::string company_;

  /** The category of the Excel Document. */
  std::string category_;

  /** The keywords of the Excel Document. */
  std::string keywords_;

  /** The comment field of the Excel Document. */
  std::string comments_;

  /** The status of the Excel Document. */
  std::string status_;

  /** The hyperlink base URL of the Excel Document. */
  ///     const char *hyperlink_base;

  /** The file creation date/time shown in Excel. This defaults to the
   * current time and date. If you wish to create files that are
   * binary equivalent (for the same input data) then you should set this
   * creation date/time to a known value. */
  std::chrono::system_clock::time_point created_;

  /** The file modification date/time shown in Excel. */
  std::chrono::system_clock::time_point modified_;
};

/**
 * @brief Workbook options.
 *
 * Optional parameters when creating a new workbook object via
 * workbook_new_opt().
 *
 * The following properties are supported:
 *
 * - `constant_memory`: This option reduces the amount of data stored in
 *   memory so that large files can be written efficiently. This option is off
 *   by default. See the notes below for limitations when this mode is on.
 *
 * - `tmpdir`: Xlsxwriter++ stores workbook data in temporary files prior to
 *   assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tmpdir` option.
 *
 * - `use_zip64`: Make the zip library use ZIP64 extensions when writing very
 *   large xlsx files to allow the zip container, or individual XML files
 *   within it, to be greater than 4 GB. See [ZIP64 on Wikipedia][zip64_wiki]
 *   for more information. This option is off by default.
 *
 *   [zip64_wiki]: https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64

 * - `output_buffer`: Output to a buffer instead of a file. The buffer must be
 *   freed manually by calling free(). This option can only be used if filename
 *   is NULL.
 *
 * - `output_buffer_size`: Used with output_buffer to get the size of the
 *   created buffer. This option can only be used if filename is NULL.
 *
 * @note In `constant_memory` mode each row of in-memory data is written to
 * disk and then freed when a new row is started via one of the
 * `worksheet_write_*()` functions. Therefore, once this option is active data
 * should be written in sequential row by row order. For this reason
 * `worksheet_merge_range()` and some other row based functionality doesn't
 * work in this mode. See @ref ww_mem_constant for more details.
 *
 * @note Also, in `constant_memory` mode the library uses temp file storage
 * for worksheet data. This can lead to an issue on OSes that map the `/tmp`
 * directory into memory since it is possible to consume the "system" memory
 * even though the "process" memory remains constant. In these cases you
 * should use an alternative temp file location by using the `tmpdir` option
 * shown above. See @ref ww_mem_temp for more details.
 */
/// typedef struct lxw_workbook_options {
/** Optimize the workbook to use constant memory for worksheets. */
///     uint8_t constant_memory;

/** Directory to use for the temporary files created by Xlsxwriter++. */
///     const char *tmpdir;

/** Allow ZIP64 extensions when creating the xlsx file zip container. */
///     uint8_t use_zip64;

/** Output buffer to use instead of writing to a file */
///     const char **output_buffer;

/** Used with output_buffer to get the size of the created buffer */
///     size_t *output_buffer_size;
/// } lxw_workbook_options;

/**
 * @brief Representation of an Excel workbook.
 */
class workbook_t
{
public:
  workbook_t();

  /**
   * @brief Validate a worksheet or chartsheet name.
   *
   * @param sheetname Sheet name to validate.
   *
   * @throw `xwpp_exception_t` or derived exception.
   *
   * This function is used to validate a worksheet or chartsheet name according
   * to the rules used by Excel:
   *
   * - The name cannot be blank.
   * - The name is less than or equal to 31 UTF-8 characters.
   * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
   * - The name doesn't start or end with an apostrophe.
   * - The name isn't already in use. (Case insensitive, see the note below).
   *
   * @code
   *     workbook.validate_sheetname("Foglio");
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
   */
  void validate_sheetname(std::string_view sheetname) const;

  /**
   * @brief Add a new worksheet to a workbook.
   *
   * @param sheetname Optional worksheet name, defaults to Sheet1, etc.
   *
   * @return A reference on Worksheet object.
   *
   * The `%add_worksheet()` function adds a new worksheet to a workbook.
   *
   * At least one worksheet should be added to a new workbook: The @ref
   * worksheet.h "Worksheet" object is used to write data and configure a
   * worksheet in the workbook.
   *
   * The `sheetname` parameter is optional. If it is not provided the default
   * Excel convention will be followed, i.e. Sheet1, Sheet2, etc.:
   *
   * @code
   *     worksheet = workbook_add_worksheet(workbook, NULL  );     // Sheet1
   *     worksheet = workbook_add_worksheet(workbook, "Foglio2");  //
  Foglio2
   *     worksheet = workbook_add_worksheet(workbook, "Data");     // Data
   *     worksheet = workbook_add_worksheet(workbook, NULL  );     // Sheet4
   *
   * @endcode
   *
   * @image html workbook02.png
   *
   * The worksheet name must be a valid Excel worksheet name, i.e:
   *
   * - The name cannot be blank.
   * - The name is less than or equal to 31 UTF-8 characters.
   * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
   * - The name doesn't start or end with an apostrophe.
   * - The name isn't already in use. (Case insensitive).
   *
   * If any of these errors are encountered the function will return NULL.
   * You can check for valid name using the `validate_sheetname()`
   * function.
   *
   * @note You should also avoid using the worksheet name "History" (case
   * insensitive) which is reserved in English language versions of
   * Excel. Non-English versions may have restrictions on the equivalent word.
   */
  worksheet_t& add_worksheet(std::string_view sheetname);
  worksheet_t& add_worksheet();

  // TODO Add builder API to create format
  /**
   * @brief Create a new @ref format.h "Format" object to formats cells in
   *        worksheets.
   *
   * @param workbook Pointer to a lxw_workbook instance.
   *
   * @return A lxw_format instance.
   *
   * The `workbook_add_format()` function can be used to create new @ref
   * format.h "Format" objects which are used to apply formatting to a cell.
   *
   * @code
   *    // Create the Format.
   *    lxw_format *format = workbook_add_format(workbook);
   *
   *    // Set some of the format properties.
   *    format_set_bold(format);
   *    format_set_font_color(format, LXW_COLOR_RED);
   *
   *    // Use the format to change the text format in a cell.
   *    worksheet_write_string(worksheet, 0, 0, "Hello", format);
   * @endcode
   *
   * See @ref format.h "the Format object" and @ref working_with_formats
   * sections for more details about Format properties and how to set them.
   *
   */
  format_t* add_format();

  [[nodiscard]] std::string assemble_xml_file();

  /**
   * @brief Set the document properties such as Title, Author etc.
   *
   * @param workbook   Pointer to a lxw_workbook instance.
   * @param properties Document properties to set.
   *
   * @return A #lxw_error.
   *
   * The `%set_properties` function can be used to set the document
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
   *     // Create a properties structure and set some of the fields.
   *     doc_properties_t properties = {
   *         .title    = "This is an example spreadsheet",
   *         .subject  = "With document properties",
   *         .author   = "John McNamara",
   *         .manager  = "Dr. Heinz Doofenshmirtz",
   *         .company  = "of Wolves",
   *         .category = "Example spreadsheets",
   *         .keywords = "Sample, Example, Properties",
   *         .comments = "Created with Xlsxwriter++",
   *         .status   = "Quo",
   *     };
   *
   *     // Set the properties in the workbook.
   *     workbook.set_properties(properties);
   * @endcode
   *
   * @image html doc_properties.png
   *
   * The `created` parameter sets the file creation date/time shown in
   * Excel. This defaults to the current time and date if set to 0. If you wish
   * to create files that are binary equivalent (for the same input data) then
   * you should set this creation date/time to a known value using a `time_t`
   * value.
   *
   */
  void set_properties(const doc_properties_t& properties);

  /**
   * @brief Set a custom document text property.
   *
   * @param workbook Pointer to a lxw_workbook instance.
   * @param name     The name of the custom property.
   * @param value    The value of the custom property.
   *
   * @return A #lxw_error.
   *
   * The `%set_custom_property()` function can be used to set one
   * or more custom document text properties not covered by the standard
   * properties in the `set_properties()` function above.
   *
   *  For example:
   *
   * @code
   *     set_custom_property(workbook, "Checked by", "Eve");
   * @endcode
   *
   * @image html custom_properties.png
   *
   * There are several `set_custom_property()` overload functions for each
   * of the custom property types supported by Excel:
   *
   * - text/string
   * - number (int32_t and double)
   * - datetime
   * - boolean
   *
   * **Note**: the name and value parameters are limited to 255 characters
   * by Excel.
   */
  void set_custom_property(std::string_view name, const std::string& value);
  void set_custom_property(std::string_view name, const char* value);
  // TODO Add overload for all integer types (template)
  void set_custom_property(std::string_view name, int32_t value);
  // TODO Add overload for all float types (template)
  void set_custom_property(std::string_view name, double value);
  void set_custom_property(std::string_view name, bool value);
  // TODO Add overload with tm, ...
  void set_custom_property(std::string_view name, const std::chrono::system_clock::time_point& value);
  void set_custom_property(std::string_view name, const std::chrono::year_month_day& value);

  /**
   * @brief Get the default URL format used with `worksheet_write_url()`.
   *
   * @param  workbook Pointer to a lxw_workbook instance.
   * @return A lxw_format instance that has hyperlink properties set.
   *
   * This function returns a lxw_format instance that is used for the default
   * blue underline hyperlink in the `worksheet_write_url()` function when a
   * format isn't specified:
   *
   * @code
   *     lxw_format *url_format = workbook_get_default_url_format(workbook);
   * @endcode
   *
   * The format is the hyperlink style defined by Excel for the default theme.
   * This format is only ever required when overwriting a string URL with
   * data of a different type. See the example below.
   */
  format_t* get_default_url_format() const;

  /**
   * @brief Saves the workbook objet in Excel file.
   *
   * @param filename The name of the Excel file to create.
   *
   * The `%workbook::save()` function writes the Excel file to disk.
   *
   * @code
   *     workbook.save(workbook);
   * @endcode
   */
  // TODO Add API with C++ filesystem (std::path)
  void save(std::string_view filename);

  /**
   * @brief Create a new chart to be added to a worksheet:
   *
   * @param workbook   Pointer to a lxw_workbook instance.
   * @param chart_type The type of chart to be created. See #lxw_chart_type.
   *
   * @return A lxw_chart object.
   *
   * The `%workbook_add_chart()` function creates a new chart object that can
   * be added to a worksheet:
   *
   * @code
   *     // Create a chart object.
   *     lxw_chart *chart = workbook_add_chart(workbook, LXW_CHART_COLUMN);
   *
   *     // Add data series to the chart.
   *     chart_add_series(chart, NULL, "Sheet1!$A$1:$A$5");
   *     chart_add_series(chart, NULL, "Sheet1!$B$1:$B$5");
   *     chart_add_series(chart, NULL, "Sheet1!$C$1:$C$5");
   *
   *     // Insert the chart into the worksheet
   *     worksheet_insert_chart(worksheet, CELL("B7"), chart);
   * @endcode
   *
   * The available chart types are defined in #lxw_chart_type. The types of
   * charts that are supported are:
   *
   * | Chart type                               | Description | |
   * :--------------------------------------- |
   * :------------------------------------  | | #LXW_CHART_AREA | Area chart. | |
   * #LXW_CHART_AREA_STACKED                  | Area chart - stacked. | |
   * #LXW_CHART_AREA_STACKED_PERCENT          | Area chart - percentage stacked. |
   * | #LXW_CHART_BAR                           | Bar chart. | |
   * #LXW_CHART_BAR_STACKED                   | Bar chart - stacked. | |
   * #LXW_CHART_BAR_STACKED_PERCENT           | Bar chart - percentage stacked. |
   * | #LXW_CHART_COLUMN                        | Column chart. | |
   * #LXW_CHART_COLUMN_STACKED                | Column chart - stacked. | |
   * #LXW_CHART_COLUMN_STACKED_PERCENT        | Column chart - percentage stacked.
   * | | #LXW_CHART_DOUGHNUT                      | Doughnut chart. | |
   * #LXW_CHART_LINE                          | Line chart. | |
   * #LXW_CHART_LINE_STACKED                  | Line chart - stacked. | |
   * #LXW_CHART_LINE_STACKED_PERCENT          | Line chart - percentage stacked. |
   * | #LXW_CHART_PIE                           | Pie chart. | |
   * #LXW_CHART_SCATTER                       | Scatter chart. | |
   * #LXW_CHART_SCATTER_STRAIGHT              | Scatter chart - straight. | |
   * #LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS | Scatter chart - straight with
   * markers. | | #LXW_CHART_SCATTER_SMOOTH                | Scatter chart -
   * smooth.                | | #LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS   | Scatter
   * chart - smooth with markers.   | | #LXW_CHART_RADAR                         |
   * Radar chart.                           | | #LXW_CHART_RADAR_WITH_MARKERS |
   * Radar chart - with markers.            | | #LXW_CHART_RADAR_FILLED | Radar
   * chart - filled.                  |
   *
   *
   *
   * See @ref chart.h for details.
   */
  chart_t& add_chart(chart_type_t chart_type);

  void unset_default_url_format();

private:
  // TODO packager_t needs to access to workbook field.
  friend class packager_t;
  friend class rich_value_t;

  // TODO Should be provided to worksheet as callback
  int32_t get_xf_index(format_t* format);

  void prepare_fonts();
  void prepare_num_formats();
  void prepare_borders();
  void prepare_workbook();

  // Iterate through the worksheets and set up the VML objects.
  void prepare_vml();

  // Iterate through the worksheets and store any defined names used for print
  // ranges or repeat rows/columns.
  void prepare_defined_names();

  // Iterate through the worksheets and set up any chart or image drawings.
  void prepare_drawings();

  // Add "cached" data to charts to provide the numCache and strCache data for
  // series and title/axis ranges.
  void add_chart_cache_data();

  // Iterate through the worksheets and set up the table objects.
  void prepare_tables();

  void prepare_fills();

  void populate_range(series_range_t& range);
  void populate_range_data_cache(series_range_t& range);
  void populate_range_dimensions(series_range_t& range);

  void store_image_type(image_types_t image_type);

  /**
   * @brief Get a worksheet object from its name.
   *
   * @param workbook Pointer to a lxw_workbook instance.
   * @param name     Worksheet name.
   *
   * @return A lxw_worksheet object.
   *
   * This function returns a lxw_worksheet object reference based on its name:
   *
   * @code
   *     worksheet = workbook_get_worksheet_by_name(workbook, "Sheet1");
   * @endcode
   *
   */
  const worksheet_t* get_worksheet_by_name(std::string_view name) const;

  [[nodiscard]] std::string write_workbook() const;
  [[nodiscard]] std::string write_file_version() const;
  [[nodiscard]] std::string write_file_sharing() const;
  [[nodiscard]] std::string write_workbook_pr() const;
  [[nodiscard]] std::string write_workbook_view() const;
  [[nodiscard]] std::string write_book_views() const;
  [[nodiscard]] std::string write_sheet(std::string_view name, uint32_t sheet_id/*, uint8_t
                         hidden*/) const;
  [[nodiscard]] std::string write_sheets() const;
  [[nodiscard]] std::string write_defined_names() const;
  [[nodiscard]] std::string write_calc_pr() const;

  static const size_t XWPP_SHEETNAME_MAX = 31;

  // Use list to not invalidate referenced owned by caller in case of insertion
  // of new sheet
  std::list<std::variant<worksheet_t>> sheets_;

  ///     struct lxw_worksheets *worksheets;
  ///     struct lxw_chartsheets *chartsheets;
  std::map<std::string, worksheet_t*> worksheet_names_;
  ///     struct lxw_chartsheet_names *chartsheet_names;
  std::map<std::string, uint32_t> image_md5_;
  std::map<std::string, uint32_t> embedded_image_md5_;
  std::map<std::string, uint32_t> header_image_md5_;
  std::map<std::string, uint32_t> background_md5_;
  // Use list to not invalidate referenced owned by caller in case of insertion
  // of new format
  std::list<chart_t> charts_;
  std::vector<chart_t*> ordered_charts_;
  // Use list to not invalidate referenced owned by caller in case of insertion
  // of new format
  std::list<format_t> formats_;

  ///     struct lxw_defined_names *defined_names;
  shared_strings_t sst_;
  doc_properties_t properties_;
  std::vector<custom_property_t> custom_properties_;

  ///     char *filename;
  ///     lxw_workbook_options options;

  uint16_t num_sheets_          = 0; // TODO Needed ?
  uint16_t num_worksheets_      = 0;
  ///     uint16_t num_chartsheets;
  ///     uint16_t first_sheet;
  uint16_t active_sheet_        = 0;
  ///     uint16_t num_xf_formats;
  ///     uint16_t num_dxf_formats;
  uint16_t num_format_count_    = 0;
  uint16_t drawing_count_       = 0;
  uint16_t comment_count_       = 0;
  uint32_t num_embedded_images_ = 0;
  uint16_t window_width_        = 16095;
  uint16_t window_height_       = 9660;

  uint16_t font_count_     = 0;
  uint16_t border_count_   = 0;
  uint16_t fill_count_     = 0;
  ///     uint8_t optimize;
  uint16_t max_url_length_ = 2079;
  ///     uint8_t read_only;

  bool has_png_                         = false;
  bool has_jpeg_                        = false;
  bool has_bmp_                         = false;
  bool has_gif_                         = false;
  bool has_vml_                         = false;
  bool has_comments_                    = false;
  bool has_metadata_                    = false;
  bool has_embedded_images_             = false;
  bool has_dynamic_functions_           = false;
  bool has_embedded_image_descriptions_ = false;

  // TODO Combine with unordered_set to optimize search and assure uniqueness.
  // TODO And encapsule this combination and related functions in a dedicated types.
  std::vector<format_t*> used_xf_formats_;

  ///     lxw_hash_table *used_dxf_formats;

  std::string vba_project_;
  std::string vba_project_signature_;
  ///     char *vba_codename;

  ///     uint8_t use_1904_epoch;

  format_t* default_url_format_;
};

/**
 * @brief Create a new workbook object, and set the workbook options.
 *
 * @param filename The name of the new Excel file to create.
 * @param options  Workbook options.
 *
 * @return A lxw_workbook instance.
 *
 * This function is the same as the `workbook_new()` constructor but allows
 * additional options to be set.
 *
 * @code
 *    lxw_workbook_options options = {.constant_memory = LXW_TRUE,
 *                                    .tmpdir = "C:\\Temp",
 *                                    .use_zip64 = LXW_FALSE,
 *                                    .output_buffer = NULL,
 *                                    .output_buffer_size = NULL};
 *
 *    lxw_workbook  *workbook  = workbook_new_opt("filename.xlsx",
&options);
 * @endcode
 *
 * The options that can be set via #lxw_workbook_options are:
 *
 * - `constant_memory`: This option reduces the amount of data stored in
 *   memory so that large files can be written efficiently. This option is off
 *   by default. See the note below for limitations when this mode is on.
 *
 * - `tmpdir`: Xlsxwriter++ stores workbook data in temporary files prior to
 *   assembling the final XLSX file. The temporary files are created in the
 *   system's temp directory. If the default temporary directory isn't
 *   accessible to your application, or doesn't contain enough space, you can
 *   specify an alternative location using the `tmpdir` option.
 *
 * - `use_zip64`: Make the zip library use ZIP64 extensions when writing very
 *   large xlsx files to allow the zip container, or individual XML files
 *   within it, to be greater than 4 GB. See [ZIP64 on Wikipedia][zip64_wiki]
 *   for more information. This option is off by default.
 *
 *   [zip64_wiki]: https://en.wikipedia.org/wiki/Zip_(file_format)#ZIP64
 *
 * - `output_buffer`: Output to a memory buffer instead of a file. The buffer
 *   must be freed manually by calling `free()`. This option can only be used if
 *   filename is NULL.
 *
 * - `output_buffer_size`: Used with output_buffer to get the size of the
 *   created buffer. This option can only be used if filename is `NULL`.
 *
 * @note In `constant_memory` mode each row of in-memory data is written to
 * disk and then freed when a new row is started via one of the
 * `worksheet_write_*()` functions. Therefore, once this option is active data
 * should be written in sequential row by row order. For this reason
 * `worksheet_merge_range()` and some other row based functionality doesn't
 * work in this mode. See @ref ww_mem_constant for more details.
 *
 * @note Also, in `constant_memory` mode the library uses temp file storage
 * for worksheet data. This can lead to an issue on OSes that map the `/tmp`
 * directory into memory since it is possible to consume the "system" memory
 * even though the "process" memory remains constant. In these cases you
 * should use an alternative temp file location by using the `tmpdir` option
 * shown above. See @ref ww_mem_temp for more details.
 */
/// lxw_workbook *workbook_new_opt(const char *filename,
///                                lxw_workbook_options *options);

/**
 * @brief Add a new chartsheet to a workbook.
 *
 * @param workbook  Pointer to a lxw_workbook instance.
 * @param sheetname Optional chartsheet name, defaults to Chart1, etc.
 *
 * @return A lxw_chartsheet object.
 *
 * The `%workbook_add_chartsheet()` function adds a new chartsheet to a
 * workbook. The @ref chartsheet.h "Chartsheet" object is like a worksheet
 * except it displays a chart instead of cell data.
 *
 * @image html chartsheet.png
 *
 * The `sheetname` parameter is optional. If it is `NULL` the default
 * Excel convention will be followed, i.e. Chart1, Chart2, etc.:
 *
 * @code
 *     chartsheet = workbook_add_chartsheet(workbook, NULL  );     // Chart1
 *     chartsheet = workbook_add_chartsheet(workbook, "My Chart"); // My
Chart
 *     chartsheet = workbook_add_chartsheet(workbook, NULL  );     // Chart3
 *
 * @endcode
 *
 * The chartsheet name must be a valid Excel worksheet name, i.e.:
 *
 * - The name cannot be blank.
 * - The name is less than or equal to 31 UTF-8 characters.
 * - The name doesn't contain any of the characters: ` [ ] : * ? / \ `
 * - The name doesn't start or end with an apostrophe.
 * - The name isn't already in use. (Case insensitive).
 *
 * If any of these errors are encountered the function will return NULL.
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
/// lxw_chartsheet *workbook_add_chartsheet(lxw_workbook *workbook,
///                                         const char *sheetname);

/**
 * @brief Set a custom document date or time property.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The name of the custom property.
 * @param datetime The value of the custom property.
 *
 * @return A #lxw_error.
 *
 * Set a custom date or time number property.
 * See `set_custom_property()` above for details.
 *
 * @code
 *     lxw_datetime datetime  = {2016, 12, 1,  11, 55, 0.0};
 *
 *     workbook_set_custom_property_datetime(workbook, "Date completed",
 * &datetime);
 * @endcode
 */
/// lxw_error workbook_set_custom_property_datetime(lxw_workbook *workbook,
///                                                 const char *name,
///                                                 lxw_datetime *datetime);

/**
 * @brief Create a defined name in the workbook to use as a variable.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     The defined name.
 * @param formula  The cell or range that the defined name refers to.
 *
 * @return A #lxw_error.
 *
 * This function is used to defined a name that can be used to represent a
 * value, a single cell or a range of cells in a workbook: These defined names
 * can then be used in formulas:
 *
 * @code
 *     workbook_define_name(workbook, "Exchange_rate", "=0.96");
 *     worksheet_write_formula(worksheet, 2, 1, "=Exchange_rate", NULL);
 *
 * @endcode
 *
 * @image html defined_name.png
 *
 * As in Excel a name defined like this is "global" to the workbook and can be
 * referred to from any worksheet:
 *
 * @code
 *     // Global workbook name.
 *     workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10");
 * @endcode
 *
 * It is also possible to define a local/worksheet name by prefixing it with
 * the sheet name using the syntax `'sheetname!definedname'`:
 *
 * @code
 *     // Local worksheet name.
 *     workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * If the sheet name contains spaces or special characters you must follow the
 * Excel convention and enclose it in single quotes:
 *
 * @code
 *     workbook_define_name(workbook, "'New Data'!Sales", "=Sheet2!$G$1:$G$10");
 * @endcode
 *
 * [Microsoft Office
 * documentation](https://support.microsoft.com/en-us/office/define-and-use-names-in-formulas-4d0f13ac-53b7-422e-afd2-abd7ff379c64).
 *
 */
/// lxw_error workbook_define_name(lxw_workbook *workbook, const char *name,
///                                const char *formula);

/**
 * @brief Get a chartsheet object from its name.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     chartsheet name.
 *
 * @return A lxw_chartsheet object.
 *
 * This function returns a lxw_chartsheet object reference based on its name:
 *
 * @code
 *     chartsheet = workbook_get_chartsheet_by_name(workbook, "Chart1");
 * @endcode
 *
 */
/// lxw_chartsheet *workbook_get_chartsheet_by_name(lxw_workbook *workbook,
///                                                 const char *name);

/**
 * @brief Add a vbaProject binary to the Excel workbook.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param filename The path/filename of the vbaProject.bin file.
 *
 * The `%workbook_add_vba_project()` function can be used to add macros or
 * functions to a workbook using a binary VBA project file that has been
 * extracted from an existing Excel xlsm file:
 *
 * @code
 *     workbook_add_vba_project(workbook, "vbaProject.bin");
 * @endcode
 *
 * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
 * have to be `vbaProject.bin`. Any suitable path/name for an existing VBA bin
 * file will do.
 *
 * Once you add a VBA project had been add to an Xlsxwriter++ workbook you
 * should ensure that the file extension is `.xlsm` to prevent Excel from
 * giving a warning when it opens the file:
 *
 * @code
 *     lxw_workbook *workbook = new_workbook("macro.xlsm");
 * @endcode
 *
 * See also @ref working_with_macros
 *
 * @return A #lxw_error.
 */
/// lxw_error workbook_add_vba_project(lxw_workbook *workbook,
///                                    const char *filename);

/**
 * @brief Add a vbaProject binary and a vbaProjectSignature binary to the Excel
 * workbook.
 *
 * @param workbook    Pointer to a lxw_workbook instance.
 * @param vba_project The path/filename of the vbaProject.bin file.
 * @param signature   The path/filename of the vbaProjectSignature.bin file.
 *
 * The `%workbook_add_signed_vba_project()` function can be used to add
 * digitally signed macros or functions to a workbook. The function adds a
 * binary VBA project file and a binary VBA project signature file that have
 * been extracted from an existing Excel xlsm file with digitally signed macros:
 *
 * @code
 *     workbook_add_signed_vba_project(workbook, "vbaProject.bin",
 * "vbaProjectSignature.bin");
 * @endcode
 *
 * Only one `vbaProject.bin` file can be added per workbook. The name doesn't
 * have to be `vbaProject.bin`. Any suitable path/name for an existing VBA bin
 * file will do. The same applies for `vbaProjectSignature.bin`.
 *
 * See also @ref working_with_macros
 *
 * @return A #lxw_error.
 */
/// lxw_error workbook_add_signed_vba_project(lxw_workbook *workbook,
///                                           const char *vba_project,
///                                           const char *signature);

/**
 * @brief Set the VBA name for the workbook.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param name     Name of the workbook used by VBA.
 *
 * The `workbook_set_vba_name()` function can be used to set the VBA name for
 * the workbook. This is sometimes required when a vbaProject macro included
 * via `workbook_add_vba_project()` refers to the workbook by a name other
 * than `ThisWorkbook`.
 *
 * @code
 *     workbook_set_vba_name(workbook, "MyWorkbook");
 * @endcode
 *
 * If an Excel VBA name for the workbook isn't specified then Xlsxwriter++
 * will use `ThisWorkbook`.
 *
 * See also @ref working_with_macros
 *
 * @return A #lxw_error.
 */
/// lxw_error workbook_set_vba_name(lxw_workbook *workbook, const char *name);

/**
 * @brief Add a recommendation to open the file in "read-only" mode.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * This function can be used to set the Excel "Read-only Recommended" option
 * that is available when saving a file. This presents the user of the file
 * with an option to open it in "read-only" mode. This means that any changes
 * to the file can't be saved back to the same file and must be saved to a new
 * file. It can be set as follows:
 *
 * @code
 *     workbook_read_only_recommended(workbook);
 * @endcode
 *
 * Which will raise a dialog like the following when opening the file:
 *
 * @image html read_only.png
 */
/// void workbook_read_only_recommended(lxw_workbook *workbook);

/**
 * @brief Set the workbook to use the 1904 epoch.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 *
 * The `%workbook_use_1904_epoch()` function can be used to set the workbook to
 * use the 1904 epoch instead of the default 1900 epoch.
 *
 * Excel supports two date epochs. The first based on 1900-01-01 is the default
 * for all Windows versions of Excel and for recent versions of Excel for macOS.
 * Older versions of Excel for macOS used a 1904-01-01 epoch. The 1904 epoch can
 * be set for compatibility with older versions of Excel or to work around the
 * Excel limitation of not being able to handle negative times.
 *
 * This function should be called before `worksheet_add_worksheet()`.
 *
 * @code
 *     workbook_use_1904_epoch(workbook);
 * @endcode
 *
 */
/// void workbook_use_1904_epoch(lxw_workbook *workbook);

/**
 * @brief Set the size of a workbook window.
 *
 * @param workbook Pointer to a lxw_workbook instance.
 * @param width    Width of the window in pixels.
 * @param height   Height of the window in pixels.
 *
 * Set the size of a workbook window. This is generally only useful on macOS
 * since Microsoft Windows uses the window size from the last time an Excel file
 * was opened/saved. The default size is 1073 x 644 pixels.
 *
 * The resulting pixel sizes may not exactly match the target screen and
 * resolution since it is based on the original Excel for Windows sizes. Some
 * trial and error may be required to get an exact size.
 */
/// void workbook_set_size(lxw_workbook *workbook,
///                        uint16_t width, uint16_t height);

/// void lxw_workbook_free(lxw_workbook *workbook);
/// void lxw_workbook_set_default_xf_indices(lxw_workbook *workbook);

/// STATIC void _write_defined_name(lxw_workbook *self,
///                                 lxw_defined_name *define_name);

/// STATIC lxw_error _store_defined_name(lxw_workbook *self, const char *name,
///                                      const char *app_name,
///                                      const char *formula, int16_t index,
///                                      uint8_t hidden);

}

#endif
