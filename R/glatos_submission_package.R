#' Convert glatos_workbook object to openxlsx Workbook
#'
#' Used for writing a GLATOS workbook submission package to disk.
#'
#' @param wb A \code{glatos_workbook} object; e.g., created by
#'   \code{\link{read_glatos_workbook}}.
#'
#' @param local_tzone A character string specifying the time zone to be used for
#'   the conversion from UTC to local time (e.g, "US/Eastern"). See details.
#'
#' @details Local time zone must be specified (via \code{local_tzone}) because
#'   \code{glatos_workbook} objects only contain time stamps in UTC. However,
#'   the GLATOS Data Submission Package (a *.zip archive containing a *.xlsm
#'   file) requires data in a local time zone (limited to \code{"US/Eastern"} or
#'   \code{"US/Central"}, which are actually expressed simply as
#'   \code{"Eastern"} or \code{"Central"} in a GLATOS Submission Package. Thus,
#'   for GLATOS Data Portal compatibility, only values of \code{local_tzone}
#'   supported are \code{"US/Eastern"} or \code{"US/Central"} are allowed.
#'
#' @details \code{as_glatos_workbook_xlsx} is essentially the inverse of
#'   \code{\link[glatos::read_glatos_workbook]{read_glatos_workbook}}, resulting
#'   in an object that resembles a GLATOS Workbook *.xlsm file in structure;
#'   namely six worksheets: 'Project', 'Locations', 'Proposed', 'Deployment'.
#'   Data for the 'Project' sheet are taken from \code{wb$metadata}. The
#'   'Proposed' sheet contains no data. Data for the 'Locations', 'Deployment',
#'   and 'Recovery' sheets are taken from \code{wb$receivers}. Data for the
#'   'Tagging' sheet are taken from \code{wb$animals}.
#'
#' @note Project-specific columns and elements of \code{wb} are not exported.
#'
#' @author Christopher Holbrook, \email{cholbrook@@usgs.gov}
#'
#' @return An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook} object.
#'
#' @examples
#'
#' #get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'                        "walleye_workbook.xlsm", package = "glatos")
#'
#' wb <- glatos::read_glatos_workbook(wb_file)
#'
#' wbx <- as_glatos_workbook_xlsx(wb, local_tzone = "US/Eastern")
#'
#' openxlsx::openXL(wbx)
#'
#' @export
as_glatos_workbook_xlsx <- function(wb, local_tzone){


  # Make new empty workbook
  wbx <- new_glatos_workbook_xlsx()


  # Read data from input glatos_workbook object

  # get data for Project sheet
  Project <- data.frame(x = unlist(wb[["metadata"]][c("project_code",
                                                      "principle_investigator",
                                                      "pi_email")]))


  # get data for Locations sheet
  Locations <- data.frame(
    glatos_array = wb$receivers$glatos_array,
    location_description = wb$receivers$location_description,
    water_body = wb$receivers$water_body,
    glatos_region = wb$receivers$glatos_region
  )

  Locations <- unique(Locations)

  Locations <- Locations[order(Locations$glatos_array),]


  # get data for Deployment sheet
  Deployments <- wb$receivers

  Deployments <- with(Deployments,
                      data.frame(
                        GLATOS_ARRAY = glatos_array,
                        OTN_ARRAY = otn_array,
                        STATION_NO = station_no,
                        CONSECUTIVE_DEPLOY_NO = consecutive_deploy_no,
                        MOORING_DROP_DEAD_DATE = mooring_drop_dead_date,
                        INTEND_LAT = intend_lat,
                        INTEND_LONG = intend_long,
                        OTN_MISSION_ID = otn_mission_id,
                        DEPLOY_DATE_TIME = deploy_date_time,
                        GLATOS_DEPLOY_DATE_TIME = deploy_date_time,
                        # note rep in next line correctly handles df with 0 rows
                        GLATOS_TIMEZONE = rep(gsub("US/", "", local_tzone),
                                              nrow(Deployments)),
                        DEPLOY_LAT = deploy_lat,
                        DEPLOY_LONG = deploy_long,
                        BOTTOM_DEPTH = bottom_depth,
                        RISER_LENGTH = riser_length,
                        INSTRUMENT_DEPTH = instrument_depth,
                        CHECWLK_COMPLETE_TIME = checwlk_complete_time,
                        STATUS_IN = status_in,
                        INS_MODEL_NO = ins_model_no,
                        GLATOS_INS_FREQUENCY = glatos_ins_frequency,
                        INS_SERIAL_NO = ins_serial_no,
                        RCV_MODEM_ADDRESS = rcv_modem_address,
                        SYNC_DATE_TIME = sync_date_time,
                        MEMORY_ERASED_AT_DEPLOY = memory_erased_at_deploy,
                        RCV_BATTERY_INSTALL_DATE = rcv_battery_install_date,
                        RCV_EXPECTED_BATTERY_LIFE = rcv_expected_battery_life,
                        RCV_VOLTAGE_AT_DEPLOY = rcv_voltage_at_deploy,
                        RCV_TILT_AFTER_DEPLOY = rcv_tilt_after_deploy,
                        DEPLOYED_BY = deployed_by,
                        COMMENTS = comments,
                        GLATOS_SEASONAL = glatos_seasonal,
                        GLATOS_PROJECT = glatos_project,
                        GLATOS_VPS = glatos_vps))

  #enforce local time zone for glatos_
  attr(Deployments$GLATOS_DEPLOY_DATE_TIME, "tzone") <- local_tzone


  # get data for Recovery sheet
  Recoveries <- wb$receivers[!is.na(wb$receivers$recover_date_time),]

  Recoveries <- with(Recoveries,
                     data.frame(
                       GLATOS_ARRAY = glatos_array,
                       OTN_ARRAY = otn_array,
                       STATION_NO = station_no,
                       CONSECUTIVE_DEPLOY_NO = consecutive_deploy_no,
                       OTN_MISSION_ID = otn_mission_id,
                       AR_CONFIRM = ar_confirm,
                       DATA_DOWNLOADED = data_downloaded,
                       INS_MODEL_NO = ins_model_no,
                       INS_SERIAL_NUMBER = ins_serial_no,
                       RECOVERED = recovered,
                       RECOVER_DATE_TIME = recover_date_time,
                       GLATOS_RECOVER_DATE_TIME = recover_date_time,
                       # note rep in next line correctly handles df with 0 rows
                       GLATOS_TIMEZONE = rep(gsub("US/", "", local_tzone),
                                             nrow(Recoveries)),
                       RECOVER_LAT = recover_lat,
                       RECOVER_LONG = recover_long,
                       GLATOS_PROJECT = glatos_project,
                       GLATOS_VPS = glatos_vps))

  #enforce local time zone for glatos_
  attr(Recoveries$GLATOS_RECOVER_DATE_TIME, "tzone") <- local_tzone


  # get data for Tagging sheet
  Tagging <- wb$animals

  Tagging <- with(Tagging,
                  data.frame(check.rows = FALSE,
                    ANIMAL_ID = animal_id,
                    TAG_TYPE = tag_type,
                    TAG_MANUFACTURER = tag_manufacturer,
                    TAG_MODEL = tag_model,
                    TAG_SERIAL_NUMBER = tag_serial_number,
                    TAG_ID_CODE = tag_id_code,
                    TAG_CODE_SPACE = tag_code_space,
                    TAG_IMPLANT_TYPE = tag_implant_type,
                    TAG_ACTIVATION_DATE = tag_activation_date,
                    EST_TAG_LIFE = est_tag_life,
                    TAGGER = tagger,
                    TAG_OWNER_PI = tag_owner_pi,
                    TAG_OWNER_ORGANIZATION = tag_owner_organization,
                    COMMON_NAME_E = common_name_e,
                    SCIENTIFIC_NAME = scientific_name,
                    CAPTURE_LOCATION = capture_location,
                    CAPTURE_LATITUDE = capture_latitude,
                    CAPTURE_LONGITUDE = capture_longitude,
                    WILD_OR_HATCHERY = wild_or_hatchery,
                    STOCK = stock,
                    LENGTH = length,
                    WEIGHT = weight,
                    LENGTH_TYPE = length_type,
                    AGE = age,
                    SEX = sex,
                    DNA_SAMPLE_TAKEN = dna_sample_taken,
                    TREATMENT_TYPE = treatment_type,
                    RELEASE_GROUP = release_group,
                    RELEASE_LOCATION = release_location,
                    RELEASE_LATITUDE = release_latitude,
                    RELEASE_LONGITUDE = release_longitude,
                    UTC_RELEASE_DATE_TIME = utc_release_date_time,
                    GLATOS_RELEASE_DATE_TIME = utc_release_date_time,
                    # note rep in next line correctly handles df with 0 rows
                    GLATOS_TIMEZONE = rep(gsub("US/", "", local_tzone),
                                          nrow(Tagging)),
                    CAPTURE_DEPTH = capture_depth,
                    TEMPERATURE_CHANGE = temperature_change,
                    HOLDING_TEMPERATURE = holding_temperature,
                    SURGERY_LOCATION = surgery_location,
                    DATE_OF_SURGERY = date_of_surgery,
                    SURGERY_LATITUDE = surgery_latitude,
                    SURGERY_LONGITUDE = surgery_longitude,
                    SEDATIVE = sedative,
                    SEDATIVE_CONCENTRATION = sedative_concentration,
                    ANAESTHETIC = anaesthetic,
                    BUFFER = buffer,
                    ANAESTHETIC_CONCENTRATION =
                      anaesthetic_concentration,
                    BUFFER_CONCENTRATION_IN_ANAESTHETIC =
                      buffer_concentration_in_anaesthetic,
                    ANESTHETIC_CONCENTRATION_IN_RECIRCULATION =
                      anesthetic_concentration_in_recirculation,
                    BUFFER_CONCENTRATION_IN_RECIRCULATION =
                      buffer_concentration_in_recirculation,
                    DISSOLVED_OXYGEN = dissolved_oxygen,
                    COMMENTS = comments,
                    GLATOS_PROJECT = glatos_project,
                    GLATOS_EXTERNAL_TAG_ID1 = glatos_external_tag_id1,
                    GLATOS_EXTERNAL_TAG_ID2 = glatos_external_tag_id2,
                    GLATOS_TAG_RECOVERED = glatos_tag_recovered,
                    GLATOS_CAUGHT_DATE = glatos_caught_date,
                    GLATOS_REWARD = glatos_reward))

  #enforce local time zone for glatos_
  attr(Tagging$GLATOS_RELEASE_DATE_TIME, "tzone") <- local_tzone


  #----------------------------------
  # Add data to glatos_workbook_xlsx

  # Project sheet
  openxlsx::writeData(wbx,
                      sheet = "Project",
                      x = Project,
                      startCol = 4,
                      startRow = 4,
                      colNames = FALSE,
                      rowNames = FALSE)

  # Locations sheet
  openxlsx::writeData(wbx,
                      sheet = "Locations",
                      x = Locations,
                      startCol = 1,
                      startRow = 3,
                      colNames = FALSE,
                      rowNames = FALSE)

  # Deployment sheet
  openxlsx::writeData(wbx,
                      sheet = "Deployment",
                      x = Deployments,
                      startCol = 1,
                      startRow = 3,
                      colNames = FALSE,
                      rowNames = FALSE)

  # Recovery sheet
  openxlsx::writeData(wbx,
                      sheet = "Recovery",
                      x = Recoveries,
                      startCol = 1,
                      startRow = 3,
                      colNames = FALSE,
                      rowNames = FALSE)

  # Tagging sheet
  openxlsx::writeData(wbx,
                      sheet = "Tagging",
                      x = Tagging,
                      startCol = 1,
                      startRow = 3,
                      colNames = FALSE,
                      rowNames = FALSE)


  #---------------
  # Update styles
  style_glatos_workbook_xlsx(wbx)


  return(wbx)
}


#' Internal function to initialize openxlsx Workbook from a glatos_workbook
#'
#' Called by \code{\link{as_glatos_workbook_xlsx}}.
#'
#' @author Christopher Holbrook, \email{cholbrook@@usgs.gov}
#'
#' @return An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook} object.
#'
new_glatos_workbook_xlsx <- function(){

  # Create empty workbook
  wbx <- openxlsx::createWorkbook()

  # Add sheets
  wbx_sheets <- c("Project", "Locations", "Proposed", "Deployment",
                  "Recovery", "Tagging")

  for(sheet_i in wbx_sheets) openxlsx::addWorksheet(wbx, sheet_i)


  # Define headers

  # Project sheet (note this sheet is horizontal (vs. columnar) layout)
  project_headers <- c("GLATOS Project Code:",
                       "Principal Investigator (PI):",
                       "PI E-Mail Address:")

  # Locations sheet
  locations_headers <- c("GLATOS_ARRAY",
                         "LOCATION_DESCRIPTION",
                         "WATER_BODY",
                         "GLATOS_REGION")

  # Proposed sheet
  proposed_headers <- c("OTN_REGION",
                        "GLATOS_ARRAY",
                        "OTN_ARRAY",
                        "STATION_NO",
                        "BOTTOM_DEPTH",
                        "PROPOSED_LAT",
                        "PROPOSED_LONG",
                        "PROPOSED_START_DATE",
                        "PROPOSED_END_DATE",
                        "GLATOS_INS_FREQUENCY",
                        "GLATOS_FUNDED",
                        "GLATOS_SEASONAL",
                        "GLATOS_PROJECT",
                        "GLATOS_VPS")

  # Deployment sheet
  deployment_headers <- c("GLATOS_ARRAY",
                          "OTN_ARRAY",
                          "STATION_NO",
                          "CONSECUTIVE_DEPLOY_NO",
                          "MOORING_DROP_DEAD_DATE",
                          "INTEND_LAT",
                          "INTEND_LONG",
                          "OTN_MISSION_ID",
                          "DEPLOY_DATE_TIME",
                          "GLATOS_DEPLOY_DATE_TIME",
                          "GLATOS_TIMEZONE",
                          "DEPLOY_LAT",
                          "DEPLOY_LONG",
                          "BOTTOM_DEPTH",
                          "RISER_LENGTH",
                          "INSTRUMENT_DEPTH",
                          "CHECWLK_COMPLETE_TIME",
                          "STATUS_IN",
                          "INS_MODEL_NO",
                          "GLATOS_INS_FREQUENCY",
                          "INS_SERIAL_NO",
                          "RCV_MODEM_ADDRESS",
                          "SYNC_DATE_TIME",
                          "MEMORY_ERASED_AT_DEPLOY",
                          "RCV_BATTERY_INSTALL_DATE",
                          "RCV_EXPECTED_BATTERY_LIFE",
                          "RCV_VOLTAGE_AT_DEPLOY",
                          "RCV_TILT_AFTER_DEPLOY",
                          "DEPLOYED_BY",
                          "COMMENTS",
                          "GLATOS_SEASONAL",
                          "GLATOS_PROJECT",
                          "GLATOS_VPS")

  # Recovery sheet
  recovery_headers <- c("GLATOS_ARRAY",
                        "OTN_ARRAY",
                        "STATION_NO",
                        "CONSECUTIVE_DEPLOY_NO",
                        "OTN_MISSION_ID",
                        "AR_CONFIRM",
                        "DATA_DOWNLOADED",
                        "INS_MODEL_NUMBER",
                        "INS_SERIAL_NUMBER",
                        "RECOVERED",
                        "RECOVER_DATE_TIME",
                        "GLATOS_RECOVER_DATE_TIME",
                        "GLATOS_TIMEZONE",
                        "RECOVER_LAT",
                        "RECOVER_LONG",
                        "GLATOS_PROJECT",
                        "GLATOS_VPS")

  # Tagging sheet
  tagging_headers <- c("ANIMAL_ID",
                       "TAG_TYPE",
                       "TAG_MANUFACTURER",
                       "TAG_MODEL",
                       "TAG_SERIAL_NUMBER",
                       "TAG_ID_CODE",
                       "TAG_CODE_SPACE",
                       "TAG_IMPLANT_TYPE",
                       "TAG_ACTIVATION_DATE",
                       "EST_TAG_LIFE",
                       "TAGGER",
                       "TAG_OWNER_PI",
                       "TAG_OWNER_ORGANIZATION",
                       "COMMON_NAME_E",
                       "SCIENTIFIC_NAME",
                       "CAPTURE_LOCATION",
                       "CAPTURE_LATITUDE",
                       "CAPTURE_LONGITUDE",
                       "WILD_OR_HATCHERY",
                       "STOCK",
                       "LENGTH",
                       "WEIGHT",
                       "LENGT_TYPE",
                       "AGE",
                       "SEX",
                       "DNA_SAMPLE_TAKEN",
                       "TREATMENT_TYPE",
                       "RELEASE_GROUP",
                       "RELEASE_LOCATION",
                       "RELEASE_LATITUDE",
                       "RELEASE_LONGITUDE",
                       "UTC_RELEASE_DATE_TIME",
                       "GLATOS_RELEASE_DATE_TIME",
                       "GLATOS_TIMEZONE",
                       "CAPTURE_DEPTH",
                       "TEMPERATURE_CHANGE",
                       "HOLDING_TEMPERATURE",
                       "SURGERY_LOCATION",
                       "DATE_OF_SURGERY",
                       "SURGERY_LATITUDE",
                       "SURGERY_LONGITUDE",
                       "SEDATIVE",
                       "SEDATIVE_CONCENTRATION",
                       "ANAESTHETIC",
                       "BUFFER",
                       "ANAESTHETIC_CONCENTRATION",
                       "BUFFER_CONCENTRATION_IN_ANAESTHETIC",
                       "ANESTHETIC_CONCENTRATION_IN_RECIRCULATION",
                       "BUFFER_CONCENTRATION_IN_RECIRCULATION",
                       "DISSOLVED_OXYGEN",
                       "COMMENTS",
                       "GLATOS_PROJECT",
                       "GLATOS_EXTERNAL_TAG_ID1",
                       "GLATOS_EXTERNAL_TAG_ID2",
                       "GLATOS_TAG_RECOVERED",
                       "GLATOS_CAUGHT_DATE",
                       "GLATOS_REWARD")


  # Add headers
  openxlsx::writeData(wbx,
            sheet = "Project",
            x = project_headers,
            startCol = 3,
            startRow = 4,
            colNames = FALSE)

  openxlsx::writeData(wbx,
            sheet = "Locations",
            x = as.data.frame(matrix(NA,
                                     nrow = 0,
                                     ncol =  length(locations_headers),
                                     dimnames =  list(NULL,
                                                      locations_headers))),
            startCol = 1,
            startRow = 2,
            colNames = TRUE)

  openxlsx::writeData(wbx,
            sheet = "Proposed",
            x = as.data.frame(matrix(NA,
                                     nrow = 0,
                                     ncol =  length(proposed_headers),
                                     dimnames =  list(NULL,
                                                      proposed_headers))),
            startCol = 1,
            startRow = 2,
            colNames = TRUE)

  openxlsx::writeData(wbx,
            sheet = "Deployment",
            x = as.data.frame(matrix(NA,
                                     nrow = 0,
                                     ncol =  length(deployment_headers),
                                     dimnames =  list(NULL,
                                                      deployment_headers))),
            startCol = 1,
            startRow = 2,
            colNames = TRUE)

  openxlsx::writeData(wbx,
            sheet = "Recovery",
            x = as.data.frame(matrix(NA,
                                     nrow = 0,
                                     ncol =  length(recovery_headers),
                                     dimnames =  list(NULL,
                                                      recovery_headers))),
            startCol = 1,
            startRow = 2,
            colNames = TRUE)

  openxlsx::writeData(wbx,
            sheet = "Tagging",
            x = as.data.frame(matrix(NA,
                                     nrow = 0,
                                     ncol =  length(tagging_headers),
                                     dimnames =  list(NULL,
                                                      tagging_headers))),
            startCol = 1,
            startRow = 2,
            colNames = TRUE)


  # Add comments

  add_comments_glatos_workbook_xlsx(wbx)


  # Set row heights and column widths

  # function to set row heights
  set_xl_row_heights <- function(wb, rh){

    for(i in 1:length(rh)) openxlsx::setRowHeights(wb,
                                                   sheet = names(rh)[i],
                                                   rows = names(rh[[i]]),
                                                   heights = rh[[i]])
  }


  # function to set column widths
  set_xl_column_widths <- function(wb, cw){

    for(i in 1:length(cw)) openxlsx::setColWidths(wb,
                                                  sheet = names(cw)[i],
                                                  cols = names(cw[[i]]),
                                                  widths = cw[[i]])
  }

  # define row heights
  row_heights <- list(
    Project =    c(`1` = 39.8,
                   `2` = 16.5,
                   `3` = 15.8,
                   `7` = 15.0),
    Locations =  c(`1` = 30),
    Proposed =   c(`1` = 54.9),
    Deployment = c(`1` = 54.9),
    Recovery =   c(`1` = 54.9),
    Tagging =    c(`1` = 54.9))


  # define column widths
  column_widths <- list(
    Project =    c(`1` = 2,
                   `2` = 2,
                   `3` = 26,
                   `4` = 100,
                   `5` = 2,
                   `6` = 2),
    Locations =   c(`1` = 14.14,
                   `2` = 29,
                   `3` = 18.29,
                   `4` = 15.14),
    Proposed =   c(`1` = 12,
                   `2` = 14.14,
                   `3` = 11,
                   `4` = 11.86,
                   `5` = 14.86,
                   `6` = 14.14,
                   `7` = 16,
                   `8` = 22.29,
                   `9` = 20.43,
                   `10` = 20.43,
                   `11` = 15.57,
                   `12` = 17.71,
                   `13` = 15.86,
                   `14` = 11.57),
    Deployment = c(`1` = 14.14,
                   `2` = 11.00,
                   `3` = 11.86,
                   `4` = 24.57,
                   `5` = 27.29,
                   `6` = 11.14,
                   `7` = 13.14,
                   `8` = 15.86,
                   `9` = 18.14,
                   `10` = 26.29,
                   `11` = 17.57,
                   `12` = 11.14,
                   `13` = 13.14,
                   `14` = 14.86,
                   `15` = 13.14,
                   `16` = 18.86,
                   `17` = 22.00,
                   `18` = 9.860,
                   `19` = 14.57,
                   `20` = 23.00,
                   `21` = 14.14,
                   `22` = 21.43,
                   `23` = 15.86,
                   `24` = 27.57,
                   `25` = 26.57,
                   `26` = 27.14,
                   `27` = 24.43,
                   `28` = 22.86,
                   `29` = 12.57,
                   `30` = 10.71,
                   `31` = 17.71,
                   `32` = 15.86,
                   `33` = 11.57,
                   `35` = 9.000,
                   `37` = 15.14,
                   `41` = 10.29,
                   `45` = 14.86,
                   `50` = 15.00,
                   `51` = 10.00),
    Recovery =   c(`1` = 14.14,
                   `2` = 11,
                   `3` = 11.86,
                   `4` = 24.57,
                   `5` = 15.86,
                   `6` = 12.14,
                   `7` = 19.57,
                   `8` = 19.57,
                   `9` = 19.14,
                   `10` = 10.71,
                   `11` = 19.57,
                   `12` = 27.71,
                   `13` = 17.57,
                   `14` = 12.71,
                   `15` = 14.57,
                   `16` = 15.86,
                   `17` = 11.57,
                   `18` = 13.29,
                   `19` = 9,
                   `22` = 9,
                   `28` = 10.86,
                   `29` = 13.14),
    Tagging =    c(`1` = 13.17,
                   `2` = 8.954,
                   `3` = 19.84,
                   `4` = 11.17,
                   `5` = 19.73,
                   `6` = 12.73,
                   `7` = 16.62,
                   `8` = 18.29,
                   `9` = 21.84,
                   `10` = 12.39,
                   `11` = 7.29,
                   `12` = 14.62,
                   `13` = 27.17,
                   `14` = 17.84,
                   `15` = 16.29,
                   `16` = 18.73,
                   `17` = 17.84,
                   `18` = 19.73,
                   `19` = 18.84,
                   `20` = 5.954,
                   `21` = 7.176,
                   `22` = 7.29,
                   `23` = 12.39,
                   `24` = 3.844,
                   `25` = 3.399,
                   `26` = 19.39,
                   `27` = 16.17,
                   `28` = 14.95,
                   `29` = 17.95,
                   `30` = 17.17,
                   `31` = 18.95,
                   `32` = 23.29,
                   `33` = 26.84,
                   `34` = 17.62,
                   `35` = 15.17,
                   `36` = 21.84,
                   `37` = 22.62,
                   `38` = 18.62,
                   `39` = 17.39,
                   `40` = 17.73,
                   `41` = 19.62,
                   `42` = 8.731,
                   `43` = 25.73,
                   `44` = 12.39,
                   `45` = 6.844,
                   `46` = 29.39,
                   `47` = 40.29,
                   `48` = 46.73,
                   `49` = 42.39,
                   `50` = 18.62,
                   `51` = 10.73,
                   `52` = 15.84,
                   `53` = 25.95,
                   `54` = 25.95,
                   `55` = 25.95,
                   `56` = 25.95,
                   `57` = 15.95,
                   `58` = 21.29,
                   `59` = 25.29,
                   `60` = 33.62,
                   `61` = 33.62,
                   `62` = 33.62,
                   `63` = 12.95,
                   `64` = 15.73,
                   `65` = 15.73,
                   `66` = 12.95,
                   `67` = 19.62,
                   `68` = 19.62,
                   `69` = 19.62,
                   `70` = 19.62,
                   `71` = 19.62,
                   `72` = 19.62)
  )

  #set row heights
  set_xl_row_heights(wbx, row_heights)

  #set column widths
  set_xl_column_widths(wbx, column_widths)


  #add styles
  wbx <- style_glatos_workbook_xlsx(wbx)

  return(wbx)
}


#' Internal function to style glatos_workbook in openxlsx Workbook format
#'
#' Called by \code{\link{as_glatos_workbook_xlsx}} and internal function
#' \code{new_glatos_workbook_xlsx}.
#'
#' @param wbx An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook}
#'   object; created by \code{\link{as_glatos_workbook_xlsx}} or internal
#'   function \code{new_glatos_workbook_xlsx}.
#'
#' @author Christopher Holbrook, \email{cholbrook@@usgs.gov}
#'
#' @return An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook} object.
#'
style_glatos_workbook_xlsx <- function(wbx){

  # Define styles

  # fill required column headers green
  style_req_col_fill <- openxlsx::createStyle(fgFill = "#C4D79B")

  # set top-bottom header border
  style_header_border_topbot <- openxlsx::createStyle(
                                  border = c("top", "bottom"),
                                  borderColour = "black")

  # fill project cells grey
  style_fill_grey <- openxlsx::createStyle(fgFill = "#D9D9D9")

  # bold text
  style_bold_text <- openxlsx::createStyle(textDecoration = "bold")

  # top border medium
  style_top_border_med <- openxlsx::createStyle(border = "top",
                                                borderStyle = "medium")

  # bottom border medium
  style_bot_border_med <- openxlsx::createStyle(border = "bottom",
                                                borderStyle = "medium")

  # left border medium
  style_left_border_med <- openxlsx::createStyle(border = "left",
                                                 borderStyle = "medium")

  # right border medium
  style_right_border_med <- openxlsx::createStyle(border = "right",
                                                  borderStyle = "medium")

  # all border thin
  style_all_border_thin <- openxlsx::createStyle(border =
                                         c("top", "left", "bottom", "right"),
                                         borderStyle = "thin")

  # set project specific left double border
  style_projspec_border <- openxlsx::createStyle(border = "left",
                                                 borderStyle = "double",
                                                 borderColour = "black")

  # style otn timestamps
  style_otn_timestamp <- openxlsx::createStyle(numFmt = "yyyy-mm-ddThh:mm:ss")

  # style local (GLATOS) timestamps
  style_local_timestamp <- openxlsx::createStyle(numFmt = "yyyy-mm-dd hh:mm")

  # style date
  style_date <- openxlsx::createStyle(numFmt = "yyyy-mm-dd")


  # Apply styles

  # fill project cells grey
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_fill_grey,
                     rows = c(3:7, 3:7, 3, 7, 3:7),
                     cols = c(rep(2, 5), rep(3, 5), rep(4, 2), rep(5, 5)),
                     gridExpand = FALSE,
                     stack = TRUE)

  # bold text
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_bold_text,
                     rows = 4:6,
                     cols = 3,
                     gridExpand = FALSE,
                     stack = TRUE)

  # top border medium
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_top_border_med,
                     rows = 3,
                     cols = 2:5,
                     gridExpand = FALSE,
                     stack = TRUE)

  # bottom border medium
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_bot_border_med,
                     rows = 7,
                     cols = 2:5,
                     gridExpand = FALSE,
                     stack = TRUE)

  # left border medium
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_left_border_med,
                     rows = 3:7,
                     cols = 2,
                     gridExpand = TRUE,
                     stack = TRUE)

  # right border medium
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_right_border_med,
                     rows = 3:7,
                     cols = 5,
                     gridExpand = TRUE,
                     stack = TRUE)

  # all border thin
  openxlsx::addStyle(wbx,
                     sheet = "Project",
                     style = style_all_border_thin,
                     rows = 4:6,
                     cols = 4,
                     gridExpand = TRUE,
                     stack = TRUE)


  # set required column cells green
  openxlsx::addStyle(wbx,
                     sheet = "Locations",
                     style = style_req_col_fill,
                     rows = 2,
                     cols = 1:4,
                     gridExpand = FALSE,
                     stack = TRUE)

  # set top-bottom header border
  openxlsx::addStyle(wbx,
                     sheet = "Locations",
                     style = style_header_border_topbot,
                     rows = 2,
                     cols = 1:4,
                     gridExpand = FALSE,
                     stack = TRUE)

  # set required column cells green
  openxlsx::addStyle(wbx,
                     sheet = "Proposed",
                     style = style_req_col_fill,
                     rows = 2,
                     cols = c(2, 4, 6:14),
                     gridExpand = FALSE)

  # set top-bottom header border
  Proposed <- openxlsx::readWorkbook(wbx, sheet = "Proposed")

  openxlsx::addStyle(wbx,
                     sheet = "Proposed",
                     style = style_header_border_topbot,
                     rows = 2,
                     cols = 1:(ncol(Proposed) + 100),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set project specific left double border
  openxlsx::addStyle(wbx,
                     sheet = "Proposed",
                     style = style_projspec_border,
                     rows = 1:(nrow(Proposed) + 100),
                     cols = "O",
                     gridExpand = FALSE,
                     stack = TRUE)

  # set required column cells green
  openxlsx::addStyle(wbx,
                     sheet = "Deployment",
                     style = style_req_col_fill,
                     rows = 2,
                     cols = c(1, 3:4, 10:13, 19:21, 31:33),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set top-bottom header border
  Deployment <- openxlsx::readWorkbook(wbx, sheet = "Deployment")

  openxlsx::addStyle(wbx,
                     sheet = "Deployment",
                     style = style_header_border_topbot,
                     rows = 2,
                     cols = 1:(ncol(Deployment) + 100),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set project specific left double border
  openxlsx::addStyle(wbx,
                     sheet = "Deployment",
                     style = style_projspec_border,
                     rows = 1:(nrow(Deployment) + 100),
                     cols = "AH",
                     gridExpand = FALSE,
                     stack = TRUE)


  # style DEPLOY_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Deployment",
                     style = style_otn_timestamp,
                     rows = 1:(nrow(Deployment) + 100),
                     cols = "I",
                     stack = TRUE)

  # style GLATOS_DEPLOY_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Deployment",
                     style = style_local_timestamp,
                     rows = 1:(nrow(Deployment) + 100),
                     cols = "J",
                     stack = TRUE)


  # set required column cells green
  openxlsx::addStyle(wbx,
                     sheet = "Recovery",
                     style = style_req_col_fill,
                     rows = 2,
                     cols = c(1, 3:4, 9, 12:13, 16:17),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set top-bottom header border
  Recovery <- openxlsx::readWorkbook(wbx, sheet = "Recovery")

  openxlsx::addStyle(wbx,
                     sheet = "Recovery",
                     style = style_header_border_topbot,
                     rows = 2,
                     cols = 1:(ncol(Recovery) + 100),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set project specific left double border
  openxlsx::addStyle(wbx,
                     sheet = "Recovery",
                     style = style_projspec_border,
                     rows = 1:(nrow(Recovery) + 100),
                     cols = "R",
                     gridExpand = FALSE,
                     stack = TRUE)

  # style RECOVER_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Recovery",
                     style = style_otn_timestamp,
                     rows = 1:(nrow(Recovery) + 100),
                     cols = "K",
                     stack = TRUE)

  # style GLATOS_DEPLOY_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Recovery",
                     style = style_local_timestamp,
                     rows = 1:(nrow(Recovery) + 100),
                     cols = "L",
                     stack = TRUE)

  # set required column cells green
  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_req_col_fill,
                     rows = 2,
                     cols = c(6:7, 14:15, 29, 33:34, 52, 55:56),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set top-bottom header border
  Tagging <- openxlsx::readWorkbook(wbx, sheet = "Tagging")


  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_header_border_topbot,
                     rows = 2,
                     cols = 1:(ncol(Tagging) + 100),
                     gridExpand = FALSE,
                     stack = TRUE)

  # set project specific left double border
  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_projspec_border,
                     rows = 1:(nrow(Recovery) + 100),
                     cols = "BF",
                     gridExpand = FALSE,
                     stack = TRUE)

  # style DATE
  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_date,
                     rows = 1:(nrow(Tagging) + 100),
                     cols = c("I", "AM", "BD"),
                     gridExpand = TRUE,
                     stack = TRUE)

  # style RELEASE_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_otn_timestamp,
                     rows = 1:(nrow(Tagging) + 100),
                     cols = "AF",
                     stack = TRUE)

  # style GLATOS_RELEASE_DATE_TIME
  openxlsx::addStyle(wbx,
                     sheet = "Tagging",
                     style = style_local_timestamp,
                     rows = 1:(nrow(Tagging) + 100),
                     cols = "AG",
                     stack = TRUE)
 return(wbx)

}


#' Add comments to a glatos_workbook in openxlsx Workbook format
#'
#' Internal function called by \code{\link{as_glatos_workbook_xlsx}}.
#'
#' @param wbx An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook}
#'   object; created by \code{\link{as_glatos_workbook_xlsx}} or internal
#'   function \code{new_glatos_workbook_xlsx}.
#'
#' @author Christopher Holbrook, \email{cholbrook@@usgs.gov}
#'
#' @details Content for comments are read from GLATOS Data Dictionary:
#'   \code{data\{GLATOSWeb_data_dictionary\}}.
#'
#' @return An \code{\link[openxlsx::openxlsx]{openxlsx}} \code{Workbook} object.
#'
add_comments_glatos_workbook_xlsx <- function(wbx){

  # Get comments from example data
  data("GLATOSWeb_data_dictionary")

  dd <- GLATOSWeb_data_dictionary

  wbx_sheets <- unique(dd$sheet)

  # Add comments (from input file)
  for(i in 1:length(wbx_sheets)){

    dd_i <- dd[dd$sheet == wbx_sheets[i], ]

    txt_i <- with(dd_i, paste0(
                  field_name,
                  "\n\nExample: ", example,
                  "\n\nFormat: ", format_units))

    # skip Project sheet
    if(wbx_sheets[i] == "Project") next

    for(j in 1:nrow(dd_i)) {

      # Define style
      st1 <- openxlsx::createStyle("Tahoma", fontSize = 8)


      comment_ij <- openxlsx::createComment(
                      comment = txt_i[j],
                      author = "GLATOS",
                      style = st1,
                      visible = FALSE)

      openxlsx::writeComment(wbx,
                             sheet = wbx_sheets[i],
                             col = dd_i$field_number[j],
                             row = 2,
                             comment = comment_ij)

    } # end j
  } # end i

  return(wbx)
}


#' Write glatos_workbook object as openxlsx Workbook to *.xlsx or *.xlsm file
#'
#' Used for writing a GLATOS workbook submission package to disk.
#'
#' @param wb A \code{glatos_workbook} object; e.g., created by
#'   \code{\link{read_glatos_workbook}}.
#'
#' @param out_dir Output directory for xlsx file and zipped archive that are
#'   written to disk. If \code{NULL} (default value, the the object is written
#'   to the working directory.
#'
#' @param xl_file Optional name of output file. If \code{NULL} (default value),
#'   then standard GLATOS Submission Package naming convention is used, which
#'   consists of project code and date of file creation. E.g.,
#'   \code{"HECWL_GLATOS_20220925.xlsx"}.
#'
#' @param xl_format Format of output file. Currently only \code{"xlsx"} (default
#'   value) is supported.
#'
#' @param overwrite Should existing file with same name be overwritten?
#'
#' @author Christopher Holbrook, \email{cholbrook@@usgs.gov}
#'
#' @return A character string containing the full path to file written to disk.
#'
#' @examples
#'
#' #get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'                        "walleye_workbook.xlsm", package = "glatos")
#'
#' wb <- read_glatos_workbook(wb_file)
#'
#' wbx <- as_glatos_workbook_xlsx(wb, local_tzone = "US/Eastern")
#'
#' #openxlsx::openXL(wbx)
#'
#' write_glatos_workbook_xl(wbx)
#' @export
#'
write_glatos_workbook_xl <- function(wb,
                                     out_dir = NULL,
                                     xl_file = NULL,
                                     xl_format = "xlsx",
                                     overwrite = FALSE){

  xl_format <- match.arg(xl_format)

  # set output director if not specified
  if(is.null(out_dir)) out_dir <- getwd()

  # write to disk
  # if xl_file = NULL, use automatic naming convention consistent with macro
  if(is.null(xl_file)){

    # get GLATOS project code from Project sheet
    glatos_project <- as.character(openxlsx::readWorkbook(wb, "Project",
                                                          colNames = FALSE,
                                                          rows = 4,
                                                          cols = 4)[1])
    # make date label
    today <- format(Sys.Date(), "%Y%m%d")

    xl_file <- paste0(glatos_project, "_GLATOS_", today, ".", xl_format)
  }

  xl_file <- file.path(out_dir, xl_file)

  if(xl_format == "xlsx"){
    openxlsx::saveWorkbook(wb, xl_file, overwrite = overwrite)
  } else if (format == "xlsm") {

  }

  return(normalizePath(xl_file))

}


#' Write GLATOS Data Submission Package
#'
#' @param wb A \code{glatos_workbook} object; e.g., created by
#'   \code{\link{read_glatos_workbook}}.
#'
#' @param out_dir Output directory for xlsx file and zipped archive that are
#'   written to disk. If \code{NULL} (default value, the the object is written
#'   to the working directory.
#'
#' @param xl_file Optional name of output file. If \code{NULL} (default value),
#'   then standard GLATOS Submission Package naming convention is used, which
#'   consists of project code and date of file creation. E.g.,
#'   \code{"HECWL_GLATOS_20220925.xlsx"}.
#'
#' @param xl_format Format of output file. Currently only \code{"xlsx"} (default
#'   value) is supported.
#'
#' @param overwrite Should existing file with same name be overwritten?
#'
#' @param local_tzone A character string specifying the time zone to be used for
#'   the conversion from UTC to local time (e.g, "US/Eastern"). Passed to
#'   \code{as_glatos_workbook_xlsx}.
#'
#' @return A character string containing path and name of output file. An xlsx
#'   file and zip archive containing a copy of the xlsx file and csv versions
#'   of each sheet in the xlsx file are written to disk.
#'
#' @examples
#'
#' #get path to example GLATOS Data Workbook
#' wb_file <- system.file("extdata",
#'                        "walleye_workbook.xlsm", package = "glatos")
#'
#' wb <- glatos::read_glatos_workbook(wb_file)
#'
#' write_glatos_submission_package(wb, local_tzone = "US/Eastern",
#'                                 overwrite = TRUE)
#'
#' @export
write_glatos_submission_package <- function(wb,
                                            out_dir = NULL,
                                            xl_file = NULL,
                                            xl_format = "xlsx",
                                            overwrite = FALSE,
                                            local_tzone = NULL){

  stopifnot(inherits(wb, c("glatos_workbook", "Workbook")))


  # Set ouput dir to wd if null
  if(is.null(out_dir)) out_dir <- getwd()


  # Convert from glatos_work to openxlsx Workbook if needed
  if(inherits(wb, "glatos_workbook")){

    stopifnot("'local_tzone' must be specified." = !is.null(local_tzone))

    wbx <- as_glatos_workbook_xlsx(wb, local_tzone)
  }

  # Write xl file to disk
  xl_file <- write_glatos_workbook_xl(wbx, out_dir = out_dir,
                                       xl_file = xl_file,
                                       xl_format = xl_format,
                                       overwrite = overwrite)

  # Write csvs to disk
  temp_dir <- tempdir()

  # Function to quote character strings that contain comments
  # to mimic structure of GLATOS Submission Package csv files
  # determined by excel saveas macro
  quote_if_contains_comma <- function(x) {
    k <- grepl(",", x)
    if(any(k)) x[k] <- paste0('"', x[k], '"')
    return(x)
  }

  # Project sheet
  Project <- openxlsx::readWorkbook(wbx, "Project", colNames = FALSE)
  Project_file <- file.path(temp_dir,
                            paste0(Project[1, 2], "_GLATOS_Project.csv"))
  write.table(Project, Project_file, quote = FALSE, sep = ",",
              row.names = FALSE, col.names = FALSE)

  #Locations sheet
  Locations <- openxlsx::readWorkbook(wbx, "Locations", colNames = TRUE)
  Locations[] <- lapply(Locations, quote_if_contains_comma)
  Locations_file <- file.path(temp_dir,
                            paste0(Project[1, 2], "_GLATOS_Locations.csv"))
  write.csv(Locations, Locations_file, quote = FALSE, row.names = FALSE,
            na = "")

  #Proposed sheet
  Proposed <- openxlsx::readWorkbook(wbx, "Proposed", colNames = TRUE)
  Proposed[] <- lapply(Proposed, quote_if_contains_comma)
  Proposed_file <- file.path(temp_dir,
                              paste0(Project[1, 2], "_GLATOS_Proposed.csv"))
  write.csv(Proposed, Proposed_file, quote = FALSE, row.names = FALSE,
            na = "")

  #Deployment sheet
  Deployment <- openxlsx::readWorkbook(wbx, "Deployment", colNames = TRUE)
  Deployment[] <- lapply(Deployment, quote_if_contains_comma)

  # format timestamps
  Deployment$DEPLOY_DATE_TIME <-
    openxlsx::convertToDateTime(Deployment$DEPLOY_DATE_TIME)
  Deployment$DEPLOY_DATE_TIME <-
    format(Deployment$DEPLOY_DATE_TIME, "%Y-%m-%dT%H:%M:%S")

  Deployment$GLATOS_DEPLOY_DATE_TIME <-
    openxlsx::convertToDateTime(Deployment$GLATOS_DEPLOY_DATE_TIME)
  Deployment$GLATOS_DEPLOY_DATE_TIME <-
    format(Deployment$GLATOS_DEPLOY_DATE_TIME, "%Y-%m-%d %H:%M")

  Deployment_file <- file.path(temp_dir,
                             paste0(Project[1, 2], "_GLATOS_Deployment.csv"))

  # format geographic locations
  # GLATOSWeb requires four sig. digits to right of . in lat and lon
  Deployment$DEPLOY_LAT <- format(Deployment$DEPLOY_LAT, nsmall = 4)
  Deployment$DEPLOY_LONG <- format(Deployment$DEPLOY_LONG, nsmall = 4)

  write.csv(Deployment, Deployment_file, quote = FALSE, row.names = FALSE,
            na = "")

  #Recovery sheet
  Recovery <- openxlsx::readWorkbook(wbx, "Recovery", colNames = TRUE)
  Recovery[] <- lapply(Recovery, quote_if_contains_comma)
  Recovery_file <- file.path(temp_dir,
                               paste0(Project[1, 2], "_GLATOS_Recovery.csv"))

  # format timestamps
  Recovery$RECOVER_DATE_TIME <-
    openxlsx::convertToDateTime(Recovery$RECOVER_DATE_TIME)
  Recovery$RECOVER_DATE_TIME <-
    format(Recovery$RECOVER_DATE_TIME, "%Y-%m-%dT%H:%M:%S")

  Recovery$GLATOS_RECOVER_DATE_TIME <-
    openxlsx::convertToDateTime(Recovery$GLATOS_RECOVER_DATE_TIME)
  Recovery$GLATOS_RECOVER_DATE_TIME <-
    format(Recovery$GLATOS_RECOVER_DATE_TIME, "%Y-%m-%d %H:%M")

  write.csv(Recovery, Recovery_file, quote = FALSE, row.names = FALSE,
            na = "")

  #Tagging sheet
  Tagging <- openxlsx::readWorkbook(wbx, "Tagging", colNames = TRUE)
  Tagging[] <- lapply(Tagging, quote_if_contains_comma)
  Tagging_file <- file.path(temp_dir,
                             paste0(Project[1, 2], "_GLATOS_Tagging.csv"))

  # format timestamps
  Tagging$UTC_RELEASE_DATE_TIME <-
    openxlsx::convertToDateTime(Tagging$UTC_RELEASE_DATE_TIME)
  Tagging$UTC_RELEASE_DATE_TIME <-
    format(Tagging$UTC_RELEASE_DATE_TIME, "%Y-%m-%dT%H:%M:%S")

  Tagging$GLATOS_RELEASE_DATE_TIME <-
    openxlsx::convertToDateTime(Tagging$GLATOS_RELEASE_DATE_TIME)
  Tagging$GLATOS_RELEASE_DATE_TIME <-
    format(Tagging$GLATOS_RELEASE_DATE_TIME, "%Y-%m-%d %H:%M")


  # format dates
  Tagging$TAG_ACTIVATION_DATE <-
    openxlsx::convertToDate(Tagging$TAG_ACTIVATION_DATE)
  Tagging$TAG_ACTIVATION_DATE <-
    format(Tagging$TAG_ACTIVATION_DATE, "%Y-%m-%d")

  Tagging$DATE_OF_SURGERY <-
    openxlsx::convertToDate(Tagging$DATE_OF_SURGERY)
  Tagging$DATE_OF_SURGERY <-
    format(Tagging$DATE_OF_SURGERY, "%Y-%m-%d")

  Tagging$GLATOS_CAUGHT_DATE <-
    openxlsx::convertToDate(Tagging$GLATOS_CAUGHT_DATE)
  Tagging$GLATOS_CAUGHT_DATE <-
    format(Tagging$GLATOS_CAUGHT_DATE, "%Y-%m-%d")


  write.csv(Tagging, Tagging_file, quote = FALSE, row.names = FALSE,
            na = "")


  # Make zip archive
  zip_file <- gsub(paste0(xl_format, "$"), "zip", xl_file)

  zip(zip_file,
      files = c(xl_file,
                          Project_file,
                          Locations_file,
                          Proposed_file,
                          Deployment_file,
                          Recovery_file,
                          Tagging_file),
      extras = "-j")

  return(normalizePath(zip_file))

}


#' Convert XLSX to XLSM file
#'
#' @note I am  sure there is any need for this function.
xlsx_to_xlsm <- function(xlsx_file){

  stopifnot("input 'xlsx_file' must have extention 'xlsx'" =
              tools::file_ext(xlsx_file) == "xlsx")

  xlsx_file <- normalizePath(xlsx_file)
  xlsm_file <- gsub("xlsx$", "xlsm", xlsx_file)

  wscript_file <- tempfile(fileext = ".vbs")

  wscript <- c(
    "Set objexcel = CreateObject(\"Excel.Application\")",
    "Set objWorkbook = objexcel.Workbooks.Open(\"_xlsx_file_\")",
    "\'Objexcel.Application.Visible = True",
    "Objexcel.DisplayAlerts = false",
    "objWorkbook.SaveAs \"_xlsm_file_\", 52",
    "objWorkbook.Close",
    "Objexcel.DisplayAlerts = true",
    "Objexcel.Application.Quit",
    "\'WScript.Echo \"Finished.\"",
    "WScript.Quit")

  #update file paths
  wscript <- gsub("_xlsx_file_", xlsx_file, wscript, fixed = TRUE)
  wscript <- gsub("_xlsm_file_", xlsm_file, wscript, fixed = TRUE)

  writeLines(wscript, wscript_file)

  system_command <- paste("WScript",
                          paste0('\"', wscript_file, '\"'),
                          sep = " ")

  rx <- system2("WScript", paste0('\"', wscript_file, '\"'))

  return(xlsx_file)

}
