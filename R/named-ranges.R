#' List all the named ranges in an excel spreadsheet.
#'
#' @inheritParams read_excel
#' @export
named_ranges <- function(path) {
  path <- check_file(path)
  ext <- tolower(tools::file_ext(path))

  # Functions return a named vector
  ranges <- switch(excel_format(path),
    xls =  xls_namedranges(path),
    xlsx = xlsx_namedranges(path)
  )

  # A single named range can refer to multiple selection areas. Return only the first one
  clean.ranges <- gsub("\\+.+", "", aa$formula)

  data.frame(name = names(ranges),
             formula = ranges,
             sheet = gsub("\\!.+", "", clean.ranges),
             range = gsub(".+\\!", "", clean.ranges),
             stringsAsFactors = FALSE)

}

# Temporary functions, until xls version is implemented
xls_namedranges <- function(path) {
  stop("Named ranges are not supported for XLS files", call. = FALSE)
}
