# glatosExtras 0.2.0

## Minor changes

* Allow `glatos_workbook` object input to `write_glatos_workbook` 
  (previously requires `glatos_workbook_xlsx` object).
  
* Change name of `write_glatos_workbook_xlsx` to `write_glatos_workbook`, 
  deprecate `write_glatos_workbook_xlsx`, and (for now) simply pass all args 
  to `write_glatos_workbook`.


# glatosExtras 0.1.0

## New features

* New `read_henrys_trimble_csv()` function to read GPS data exported from Trimble in peculiar CSV format. See example file `HECWL_20220505.csv` in `inst/extdata`.


# glatosExtras 0.0.1

## New features

* New `write_glatos_submission_package()` function to write a GLATOS submission package to disk from `glatos_workbook` object. Submission package is a zip archive with GLATOS Workbook (XLSX format) and CSV version of each worksheet.

* New `write_glatos_workbook_xlsx()` function to write a GLATOS Workbook in `openxlsx Workbook` format to disk as an XLSX file.

* New `as_glatos_workbook_xlsx()` function to coerce a `glatos_workbook` object to an `openxlsx Workbook` object. 


