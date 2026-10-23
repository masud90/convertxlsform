# Tests for complex-script fonts (R/fonts.R) and fonts given to LibreOffice.

test_that("every script lists candidate fonts", {
  sc <- .cs_scripts()
  expect_setequal(names(sc), c("bengali", "arabic", "devanagari", "ethiopic",
                               "hebrew"))
  for (z in sc) {
    expect_true(length(z$fonts) >= 3L)
    expect_true(z$lo < z$hi)
  }
})

test_that(".font_installed and .font_files read the font table", {
  local_mocked_bindings(.installed_fonts = function() {
    data.frame(family = c("Noto Sans Bengali", "Arial"),
               path = c(tempfile(), tempfile()), stringsAsFactors = FALSE)
  })
  expect_equal(.font_installed(c("noto sans bengali", "Kefa", NA)),
               c(TRUE, FALSE, FALSE))
  # files that do not exist are dropped
  expect_length(.font_files("Noto Sans Bengali"), 0L)
})

test_that(".check_pdf_fonts warns when the script font is missing", {
  local_mocked_bindings(.font_files = function(family) character(0))
  m <- list(meta = list(script = "bengali",
                        font_family_cs = "Noto Sans Bengali"))
  expect_warning(.check_pdf_fonts(m), "not installed")
  m$meta$script <- NA_character_
  expect_no_warning(.check_pdf_fonts(m))
})

test_that(".check_fonts_arg expands folders and rejects missing paths", {
  d <- withr::local_tempdir()
  file.create(file.path(d, c("a.ttf", "b.OTF", "c.txt")))
  dir.create(file.path(d, "sub"))
  file.create(file.path(d, "sub", "d.ttc"))
  got <- basename(.check_fonts_arg(d))
  expect_setequal(got, c("a.ttf", "b.OTF", "d.ttc"))
  expect_equal(basename(.check_fonts_arg(file.path(d, "a.ttf"))), "a.ttf")
  expect_equal(.check_fonts_arg(NULL), character(0))
  expect_error(.check_fonts_arg(file.path(d, "none.ttf")), "not found")
  expect_error(.check_fonts_arg(1), "character vector")
})

test_that(".profile_add_fonts copies fonts into the profile", {
  d <- withr::local_tempdir()
  f <- file.path(d, "x.ttf")
  writeLines("font", f)
  prof <- file.path(d, "profile")
  expect_true(.profile_add_fonts(prof, f))
  expect_true(file.exists(file.path(prof, "user", "fonts", "1_x.ttf")))
  expect_false(.profile_add_fonts(file.path(d, "p2"), character(0)))
})

test_that("Bengali text survives in the PDF", {
  skip_on_cran()
  skip_if_not(has_libreoffice())
  skip_if(!nzchar(Sys.which("pdftotext")), "pdftotext not found")
  p <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  m <- .build_model(read_quiet(p, lang = "bn"), test_opts())
  skip_if(length(.font_files(m$meta$font_family_cs)) == 0L,
          "no Bengali font installed")
  out <- withr::local_tempfile(fileext = ".pdf")
  convertxlsform(p, output = out, lang = "bn", verbose = FALSE)
  txt <- system2("pdftotext", c("-enc", "UTF-8", shQuote(out), "-"),
                 stdout = TRUE)
  txt <- paste(enc2utf8(txt), collapse = "\n")
  # a Bengali word from the form's first section heading
  expect_match(txt, "মেটাডেটা")
})
