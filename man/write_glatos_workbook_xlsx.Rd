% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/glatos_submission_package.R
\name{write_glatos_workbook_xlsx}
\alias{write_glatos_workbook_xlsx}
\title{Write glatos_workbook object as openxlsx Workbook to *.xlsx or *.xlsm file}
\usage{
write_glatos_workbook_xlsx(
  wb,
  out_dir = NULL,
  xl_file = NULL,
  xl_format = "xlsx",
  overwrite = FALSE
)
}
\arguments{
\item{wb}{A \code{glatos_workbook} object; e.g., created by
\code{\link{read_glatos_workbook}}.}

\item{out_dir}{Output directory for xlsx file and zipped archive that are
written to disk. If \code{NULL} (default value, the the object is written
to the working directory.}

\item{xl_file}{Optional name of output file. If \code{NULL} (default value),
then standard GLATOS Submission Package naming convention is used, which
consists of project code and date of file creation. E.g.,
\code{"HECWL_GLATOS_20220925.xlsx"}.}

\item{xl_format}{Format of output file. Currently only \code{"xlsx"} (default
value) is supported.}

\item{overwrite}{Should existing file with same name be overwritten?}
}
\value{
A character string containing the full path to file written to disk.
}
\description{
Used for writing a GLATOS workbook submission package to disk.
}
\examples{

#get path to example GLATOS Data Workbook
wb_file <- system.file("extdata",
                       "walleye_workbook.xlsm", package = "glatos")

wb <- read_glatos_workbook(wb_file)

wbx <- as_glatos_workbook_xlsx(wb, local_tzone = "US/Eastern")

#openxlsx::openXL(wbx)

write_glatos_workbook_xlsx(wbx)
}
\author{
Christopher Holbrook, \email{cholbrook@usgs.gov}
}
