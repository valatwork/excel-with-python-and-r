#' Define a function called read_excel_sheets that takes two arguments: filename
#' (the name of the Excel file to be read) and single_tbl (a logical value indicating whether the
#' function should return a single table or a list of tables).
read_excel_sheets <- function(filename, single_tbl = FALSE) {
  #' Use the readxl package to extract the names of all the sheets in the Excel file specified by
  #' filename. The sheet names are stored in the sheets variable.
  sheets <- readxl::excel_sheets(filename)

  #' Starts an if statement that checks the value of the single_tbl argument.
  if (single_tbl){
    #' If single_tbl is TRUE, this line uses the purrr package’s map_df function to iterate over each
    #' sheet name in sheets and read the corresponding sheet using the read_excel function from
    #' the readxl packag
    x <- purrr::map_df(sheets, readxl::read_excel, path = filename)
  } else {
    #' If single_tbl is FALSE, the following code block is executed.
    #' the purrr package’s map function is used to iterate over each sheet name in sheets. For
    #' each sheet, the read_excel function from the readxl package is called to read the corresponding
    #' sheet from the Excel file specified by filename. The resulting DataFrame are stored in a list
    #' assigned to the x variable.
    x <- purrr::map(sheets, ~ readxl::read_excel(filename, sheet = .x))
    #' Use the set_names function from the purrr package to set the names of the elements
    #' in the x list to the sheet names in sheets.
    x <- purrr::set_names(x, sheets)
  }

  x
}

f <- "data/iris_data.xlsx"

read_excel_sheets(f, F)