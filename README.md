# glatosExtras

Niche functions that complement glatos

<br/>

### To install:


```
library(remotes)
install_github("chrisholbrook/glatosExtras")
```


<br/>

### Some things that can be done with glatosExtras:

<br/>

#### 1. Make a GLATOSWeb Submission Package from an existing GLATOS Workbook file

To make a GLATOSWeb Submission Package (a zip archive required to submit
data to the GLATOS Data Portal) from an existing workbook (xlsm or xlsx) file: 
  
  
```
library(glatos) # for read_glatos_workbook
library(glatosExtras) # for write_glatos_submission_package

# Set path to existing GLATOS Workbook file
wb_file <- "c:/user/documents/walleye_workbook.xlsm"

# Read GLATOS Workbook into R (as `glatos_workbook` object)
wb <- read_glatos_workbook(wb_file)

# Write GLATOS Submission Package (*.zip)
write_glatos_submission_package(wb, local_tzone = "US/Eastern")

```

<br/>

#### 2. Convert an existing GLATOS Workbook from macro-enabled (\*.xlsm) to macro-free (\*.xlsx) Excel file

To convert an existing \*.xlsm workbook to \*.xlsx
  
  
```
library(glatos) # for read_glatos_workbook
library(glatosExtras) # for write_glatos_workbook_xlsx

# Set path to existing GLATOS Workbook file
wb_file <- "c:/user/documents/walleye_workbook.xlsm"

# Read GLATOS Workbook into R (as `glatos_workbook` object)
wb <- read_glatos_workbook(wb_file)

# Write GLATOS Workbook (as *.xlsx file)
write_glatos_workbook(wb, local_tzone = "US/Eastern")

```

<br/>

#### 3. Make a GLATOSWeb Submission Package from a `glatos_workbook` object created from another source. 

Work in progress.




