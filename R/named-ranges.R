#' List all the named ranges in an excel spreadsheet.
#'
#' @inheritParams read_excel
#' @export
named_ranges <- function(path) {
  path <- check_file(path)
  ext <- tolower(tools::file_ext(path))

  # Functions return a named vector
  switch(excel_format(path),
    xls  = xls_namedranges(path),
    xlsx = xlsx_namedranges(path)
  )
}

# Extract correct named ranges for XLS file
xls_namedranges <- function(path) {
  x <- xls_defined_names(path)

  # Return empty data.frame if no defined names were found
  if (nrow(x) == 0L)
    return(data.frame(name = character(0), sheet = character(0), range = character(0)))

  x$range <- paste0(cellranger::num_to_letter(x$col1), x$row1,
                    ":",
                    cellranger::num_to_letter(x$col2), x$row2)
  x[, c("name", "sheet", "range")]
}

# Extract correct named ranges for XLSX file
xlsx_namedranges <- function(path) {
  ranges <- xlsx_defined_names(path)

  # Return empty data.frame if no defined names were found
  if (length(ranges) == 0L)
    return(data.frame(name = character(0), sheet = character(0), range = character(0)))

  # A single named range can refer to multiple selection areas. These are separated with '+' or ','
  valid.ranges <- grep("\\+.+", ranges, value = TRUE, invert = TRUE)

  data.frame(name = names(valid.ranges),
             sheet = gsub("\\!.+", "", valid.ranges),
             range = gsub(".+\\!", "", valid.ranges),
             stringsAsFactors = FALSE)
}
