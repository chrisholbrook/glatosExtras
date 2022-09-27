#' Data dictionary for GLATOSWeb Submission Package
#'
#' A dataset with name, example, format, units, and network requirements
#' (GLATOS, OTN) of fields (columns) in each sheet of a GLATOS Data Workbook.
#'
#' @format A data frame with 141 rows and 8 columns: \describe{
#'   \item{sheet}{Worksheet name} \item{field_number}{Column number for all
#'   sheets except Project} \item{field_name}{Column name}
#'   \item{example}{Example value} \item{format_units}{Description of format and
#'   units} \item{definition}{Definition} \item{GLATOS}{Required by GLATOS (Data
#'   Portal)?} \item{OTN}{Required by Ocean Tracking Network?} }
#'
#' @note A few corrections were made to the `Field.#` column in the original
#'   workbook before loading into this package.
#'
#' @note Used to add comments to \code{openxlsx} \code{Workbook} by
#'   \code{write_glatos_submission_package}.
#'
#' @source \url{https://www.glatos.glos.us}
"GLATOSWeb_data_dictionary"
