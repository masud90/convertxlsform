# Test fixtures: build small XLSForms on the fly.

skip_if_no_writexl <- function() {
  testthat::skip_if_not_installed("writexl")
}

# Write an xlsx with the given sheets (data.frames) and return its path.
# `survey`, `choices` and `settings` are data.frames of character columns.
# Extra sheets can be passed through `...`. Sheet names can be changed with
# `survey_name` etc. to test case-insensitivity.
make_form <- function(survey, choices = NULL, settings = NULL,
                      survey_name = "survey", choices_name = "choices",
                      settings_name = "settings", ...) {
  sheets <- list()
  sheets[[survey_name]] <- survey
  if (!is.null(choices)) sheets[[choices_name]] <- choices
  if (!is.null(settings)) sheets[[settings_name]] <- settings
  sheets <- c(sheets, list(...))
  path <- tempfile(fileext = ".xlsx")
  writexl::write_xlsx(sheets, path)
  path
}

# Shorthand data.frame of characters
fdf <- function(...) {
  data.frame(..., stringsAsFactors = FALSE, check.names = FALSE)
}

# Quiet read
read_quiet <- function(path, ...) {
  .read_xlsform(path, verbose = FALSE, ...)
}
