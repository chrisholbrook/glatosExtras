# glatosExtras
Niche functions that complement glatos


### To install:
```{r}

library(remotes)
install_github("chrisholbrook/glatosExtras")

```


### The package currently contains a set of functions with only one purpose: 

To make a GLATOSWeb Submission Package (a zip archive required to submit
 data to the GLATOS Data Portal) from an existing workbook (xlsm or xlsx) file: 
 
 
 ```{r}
 
 library(glatos)
 
 #get path to example GLATOS Data Workbook
 wb_file <- system.file("extdata",
                       "walleye_workbook.xlsm", package = "glatos")

 wb <- glatos::read_glatos_workbook(wb_file)

 write_glatos_submission_package(wb, local_tzone = "US/Eastern")

 ```
 
 Or (and perhaps more useful), from a `glatos_workbook` object created from 
 another source. 
    

