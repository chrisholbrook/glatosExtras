% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/glatos_submission_package.R
\name{write_glatos_submission_package}
\alias{write_glatos_submission_package}
\title{Write GLATOS Data Submission Package}
\usage{
write_glatos_submission_package(
  wb,
  out_dir = NULL,
  xl_file = NULL,
  xl_format = "xlsx",
  overwrite = FALSE,
  local_tzone = NULL
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

\item{local_tzone}{A character string specifying the time zone to be used for
the conversion from UTC to local time (e.g, "US/Eastern"). See details.}
}
\value{
A character string containing path and name of output file. An xlsx
file and zip archive containing a copy of the xlsx file and csv versions
of each sheet in the xlsx file are written to disk.
}
\description{
Write GLATOS Data Submission Package
}
\details{
Local time zone must be specified (via \code{local_tzone}) because
\code{glatos_workbook} objects only contain time stamps in UTC. However,
the GLATOS Data Submission Package (a *.zip archive containing a *.xlsm
file) requires data in a local time zone (limited to \code{"US/Eastern"} or
\code{"US/Central"}, which are actually expressed simply as
\code{"Eastern"} or \code{"Central"} in a GLATOS Submission Package. Thus,
for GLATOS Data Portal compatibility, only values of \code{local_tzone}
supported are \code{"US/Eastern"} or \code{"US/Central"} are allowed.
}
\examples{

#get path to example GLATOS Data Workbook
wb_file <- system.file("extdata",
                       "walleye_workbook.xlsm", package = "glatos")

wb <- glatos::read_glatos_workbook(wb_file)

write_glatos_submission_package(wb, local_tzone = "US/Eastern")

}
