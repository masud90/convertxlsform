multilang_path <- function() {
  path <- system.file("extdata/multilang/multilang.xlsx", package = "convertxlsform")
  skip_if(path == "", "multilang example not installed")
  path
}

# Language suffixes as written in the headers, e.g. "English (en)", "Bengali (bn)"
multilang_headers <- function(path) {
  cols <- names(readxl::read_excel(path, sheet = "survey"))
  sub("^label::", "", grep("^label::", cols, value = TRUE))
}

test_that("bundled multilang example loads for every language in its label columns", {
  path <- multilang_path()
  headers <- multilang_headers(path)
  expect_gt(length(headers), 0L)

  for (h in headers) {
    # Both the full header text and the bare code in parentheses should select the column
    codes <- c(h, sub(".*\\((.*)\\)$", "\\1", h))
    for (lang in unique(codes)) {
      expect_no_error(convertxlsform(path, lang = lang, verbose = FALSE))
    }
  }
})

test_that("multilang example selects the requested language's labels", {
  path <- multilang_path()
  headers <- multilang_headers(path)
  skip_if(length(headers) < 2L, "needs at least two languages")

  raw <- readxl::read_excel(path, sheet = "survey")
  res1 <- convertxlsform:::.read_and_validate_xlsform(path, lang = headers[1], verbose = FALSE)
  res2 <- convertxlsform:::.read_and_validate_xlsform(path, lang = headers[2], verbose = FALSE)
  expect_equal(res1$survey$label_selected, raw[[paste0("label::", headers[1])]])
  expect_equal(res2$survey$label_selected, raw[[paste0("label::", headers[2])]])
})

test_that("multilang example loads with the default language", {
  path <- multilang_path()
  expect_no_error(convertxlsform(path, verbose = FALSE))
})
