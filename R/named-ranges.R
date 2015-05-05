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

# Extract correct named ranges
xls_namedranges <- function(path) {
  x <- xls_defined_names(path)
  x$range <- paste0(cellranger::num_to_letter(x$col1), x$row1,
                    ":",
                    cellranger::num_to_letter(x$col2), x$row2)
  x[, c("name", "sheet", "range")]
}

# Extract correct named ranges
xlsx_namedranges <- function(path) {
  ranges <- xlsx_defined_names(path)

  # A single named range can refer to multiple selection areas. Return only the first one
  clean.ranges <- gsub("\\+.+", "", ranges)

  data.frame(name = names(ranges),
             sheet = gsub("\\!.+", "", clean.ranges),
             range = gsub(".+\\!", "", clean.ranges),
             stringsAsFactors = FALSE)
}
