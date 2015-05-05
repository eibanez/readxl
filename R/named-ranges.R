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
  ranges <- xls_defined_names(path)
  rc1 <- sprintf("R%sC%s", ranges$row1, ranges$col1)
  rc2 <- sprintf("R%sC%s", ranges$row2, ranges$col2)
  ranges$range <- paste0(cellranger::RC_to_A1(rc1), ":", cellranger::RC_to_A1(rc2))
  ranges[, c("name", "sheet", "range")]
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
