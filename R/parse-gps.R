#' Read Henry's peculiar Trimble export CSV format
#'
#' @param gps_file_in
#'
#' @details Specific to HBBS. Reads GPS data exported from a Trimble in a
#'   format like this: \cr\cr
#'
#' "Site ","MCK001","2022/05/05","09:11:26"
#' " ",45.794284940,-84.770085514,2022/05/05 13:11:44
#' " ",45.794280571,-84.770072229,2022/05/05 13:11:45
#' " ",45.794276437,-84.770057981,2022/05/05 13:11:46
#' " ",45.794272347,-84.770043991,2022/05/05 13:11:47
#' " ",45.794267438,-84.770029965,2022/05/05 13:11:48
#' "Site ","MCK002","2022/05/05","09:30:22"
#' " ",45.801236531,-84.763650723,2022/05/05 13:30:40
#' " ",45.801237444,-84.763642311,2022/05/05 13:30:41
#' " ",45.801238197,-84.763634249,2022/05/05 13:30:42
#' " ",45.801239154,-84.763625886,2022/05/05 13:30:43
#' " ",45.801240122,-84.763617542,2022/05/05 13:30:44
#'
#' @returns A data.frame with site, latitude, longitude,
#'  timestamp_utc, and source_file.
#'
#' @examples
#'
#' # Example file
#' gps_file <- system.file("inst/extdata", "HECWL_20220505.csv",
#'                         package = "glatosExtras")
#'
#' gps <- read_henrys_trimble_csv(gps_file)
#'
#'
#' \dontrun{
#' # maybe followed by something like this (for median location):
#'
#' library(dplyr)
#'
#' gps %>%
#'   group_by(site, source_file) %>%
#'   summarize(longitude = median(longitude),
#'             latitude = median(latitude),
#'             timestamp_utc = median(timestamp_utc))
#'
#' }
#'
#' returns
read_henrys_trimble_csv <- function(gps_file_in){

  gps_in <- readLines(gps_file_in)

  site_rows <- grep("^\"Site", gps_in)

  site_names <- sapply(gps_in[site_rows],
                       function(x){
                         x2 <- strsplit(x, ",")[[1]][2]
                         x2 <- gsub("\"", "", x2)
                       },
                       USE.NAMES = FALSE)

  # Create site name for each row by replicating each site name
  site_names <- rep(site_names,
                    times = diff(c(site_rows, length(gps_in) + 1)) - 1)

  # Parse lon, lat, and time; use empty first column as placeholder for site
  gps <- read.csv(text = gps_in[-site_rows],
                  header = FALSE,
                  fill = TRUE,
                  as.is = TRUE,
                  col.names = c("site", "latitude", "longitude", "timestamp_utc"))

  # Add names
  gps$site <- site_names

  # Add source file
  gps$source_file <- basename(gps_file_in)

  # Coerce timestamp to POXIXct
  gps$timestamp_utc <- as.POSIXct(gps$timestamp_utc, tz = "UTC")

  return(gps)
}
