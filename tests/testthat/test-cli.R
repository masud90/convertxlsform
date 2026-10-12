test_that("the parser maps each flag to its option", {
  p <- .parse_cli_args(c(
    "form.xlsx", "--flat-numbers", "--no-relevants", "--no-constraints",
    "--no-hints", "--no-choice-codes", "--show-names", "--include-hidden",
    "--landscape", "--overwrite", "--quiet"
  ))
  expect_equal(p$xlsform, "form.xlsx")
  expect_equal(p$options, list(
    nested_serial_numbers = FALSE, include_relevants = FALSE,
    include_constraints = FALSE, include_hints = FALSE,
    include_choice_name = FALSE, show_variable_names = TRUE,
    include_hidden = TRUE, orientation = "landscape", overwrite = TRUE,
    verbose = FALSE
  ))
})

test_that("the parser maps value options, both spellings", {
  a <- .parse_cli_args(c(
    "form.xlsx", "-o", "out.pdf", "-f", "pdf", "-l", "bn",
    "--media-root", "media", "--title", "My form", "--page-size", "letter",
    "--margins", "0.5", "--font", "Calibri", "--font-size", "12",
    "--template", "t.docx", "--max-choices", "10", "--repeat-copies", "3"
  ))
  b <- .parse_cli_args(c(
    "form.xlsx", "--output=out.pdf", "--format=pdf", "--lang=bn",
    "--media-root=media", "--title=My form", "--page-size=letter",
    "--margins=0.5", "--font=Calibri", "--font-size=12",
    "--template=t.docx", "--max-choices=10", "--repeat-copies=3"
  ))
  expect_equal(a, b)
  expect_equal(a$output, "out.pdf")
  expect_equal(a$format, "pdf")
  expect_equal(a$lang, "bn")
  expect_equal(a$options, list(
    media_root = "media", title = "My form", page_size = "letter",
    margins_in = 0.5, font_family = "Calibri", font_size = 12,
    template = "t.docx", max_choices = 10, repeat_copies = 3
  ))
})

test_that("the parser handles format both, page size case, and Inf", {
  p <- .parse_cli_args(c("f.xlsx", "-f", "both", "--page-size", "a4",
                         "--max-choices", "Inf"))
  expect_equal(p$format, c("docx", "pdf"))
  expect_equal(p$options$page_size, "A4")
  expect_equal(p$options$max_choices, Inf)
})

test_that("the position of the path does not matter", {
  p <- .parse_cli_args(c("--quiet", "-l", "fr", "form.xlsx"))
  expect_equal(p$xlsform, "form.xlsx")
  expect_equal(p$lang, "fr")
})

test_that("the parser rejects bad input", {
  bad <- function(args, msg) {
    expect_error(.parse_cli_args(args), msg, class = "cxf_cli_error")
  }
  bad(character(), "Missing")
  bad(c("f.xlsx", "--bogus"), "Unknown option")
  bad(c("f.xlsx", "-o"), "needs a value")
  bad(c("f.xlsx", "-f", "odt"), "docx, pdf or both")
  bad(c("f.xlsx", "--margins", "wide"), "needs a number")
  bad(c("f.xlsx", "--page-size", "B5"), "A4 or letter")
  bad(c("f.xlsx", "--quiet=yes"), "does not take a value")
  bad(c("a.xlsx", "b.xlsx"), "one XLSForm path")
})

test_that("help and version are parsed without a path", {
  expect_true(.parse_cli_args("--help")$help)
  expect_true(.parse_cli_args("-h")$help)
  expect_true(.parse_cli_args("--version")$version)
})

test_that("cli() prints help and version and returns 0", {
  out <- capture.output(st <- cli("--help"))
  expect_equal(st, 0L)
  expect_match(paste(out, collapse = "\n"), "--flat-numbers")
  expect_match(paste(out, collapse = "\n"), "--repeat-copies")
  out <- capture.output(st <- cli("--version"))
  expect_equal(st, 0L)
  expect_match(out, "^convertxlsform [0-9]")
})

test_that("cli() returns 1 on usage errors and prints usage to stderr", {
  err <- capture.output(st <- cli(c("--bogus", "f.xlsx")), type = "message")
  expect_equal(st, 1L)
  expect_match(err[1], "Unknown option: --bogus")
  expect_true(any(grepl("^Usage:", err)))
  err <- capture.output(st <- cli(character()), type = "message")
  expect_equal(st, 1L)
  expect_match(err[1], "Missing the XLSForm path")
})

test_that("cli() returns 1 for a missing file with a plain message", {
  stub_render_docx()
  err <- capture.output(st <- cli(c("no/such.xlsx", "--quiet")),
                        type = "message")
  expect_equal(st, 1L)
  expect_match(err[1], "^Error: XLSForm file not found")
  expect_false(any(grepl("Backtrace|Traceback", err)))
})

test_that("cli() converts a form and returns 0", {
  stub_render_docx()
  src <- copy_example()
  out <- file.path(dirname(src), "cli_out.docx")
  st <- cli(c(src, "-o", out, "--quiet"))
  expect_equal(st, 0L)
  expect_true(file.exists(out))
  # second run: file exists -> 1, with --overwrite -> 0
  capture.output(st <- cli(c(src, "-o", out, "--quiet")), type = "message")
  expect_equal(st, 1L)
  expect_equal(cli(c(src, "-o", out, "--quiet", "--overwrite")), 0L)
})

test_that("cli() passes options through to the conversion", {
  stub_render_docx()
  src <- copy_example()
  seen <- NULL
  testthat::local_mocked_bindings(
    convertxlsform = function(xlsform, output = NULL, format = NULL,
                              lang = NULL, ...) {
      seen <<- list(xlsform = xlsform, lang = lang, dots = list(...))
      invisible("x")
    }
  )
  expect_equal(cli(c(src, "-l", "bn", "--flat-numbers", "--landscape")), 0L)
  expect_equal(seen$lang, "bn")
  expect_false(seen$dots$nested_serial_numbers)
  expect_equal(seen$dots$orientation, "landscape")
})

test_that("cli() suggests --overwrite when the output exists", {
  stub_render_docx()
  src <- copy_example()
  out <- file.path(dirname(src), "cli_exists.docx")
  expect_equal(cli(c(src, "-o", out, "--quiet")), 0L)
  err <- capture.output(st <- cli(c(src, "-o", out, "--quiet")),
                        type = "message")
  expect_equal(st, 1L)
  expect_true(any(grepl("--overwrite", err, fixed = TRUE)))
  # the R function keeps its own hint
  e <- expect_error(convertxlsform(src, output = out, verbose = FALSE),
                    "overwrite = TRUE")
  expect_false(grepl("--overwrite", conditionMessage(e), fixed = TRUE))
})
