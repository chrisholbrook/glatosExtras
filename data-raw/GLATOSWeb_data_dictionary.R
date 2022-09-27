# Make a data frame containing data dictionary for the GLATOS Workbook
# and xlsx file available from the GLATOS Data Portal


# Get path to data dictionary file
dd_file <- file.path("./inst/extdata/GLATOSWeb_DataDictionaries.xlsx")


# Get sheet names
dd_sheets <- openxlsx::getSheetNames(dd_file)


# Initialize and get data

dd <- vector("list", length(dd_sheets))
names(dd) <- dd_sheets

for(i in 1:length(dd_sheets)){

  dd[[dd_sheets[i]]] <- openxlsx::readWorkbook(dd_file,
                                               sheet = dd_sheets[i],
                                               startRow = 3)

  # add sheet name
  dd[[dd_sheets[i]]]$sheet <- dd_sheets[i]


} # end i


# Combine into single data frame
dd <- do.call(rbind, dd)

row.names(dd) <- NULL

# Rename and reorder columns
dd <- dplyr::select(dd,
                    sheet,
                    field_number = `Field.#`,
                    field_name = Field.Name,ö
                    example = Example,
                    format_units = `Format/Units`,
                    definition = Definition,
                    GLATOS,
                    OTN)

# Rename object for writing
GLATOSWeb_data_dictionary <- dd


# Add to package data
usethis::use_data(GLATOSWeb_data_dictionary, overwrite = TRUE)
