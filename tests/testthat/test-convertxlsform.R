test_that("output defaults to the input folder and base name", {
  stub_render_docx()
  src <- copy_example()
  res <- convertxlsform(src, verbose = FALSE)
  expect_equal(res, file.path(dirname(normalizePath(src, winslash = "/")), "myform.docx"))
  expect_true(file.exists(res))
})

test_that("a folder output gets <basename>.<ext> inside it", {
  stub_render_docx()
  src <- copy_example()
  out <- withr::local_tempdir()
  res <- convertxlsform(src, output = out, verbose = FALSE)
  expect_equal(basename(res), "myform.docx")
  expect_equal(normalizePath(dirname(res)), normalizePath(out))
})

test_that("an explicit file path is used as given", {
  stub_render_docx()
  src <- copy_example()
  out <- file.path(withr::local_tempdir(), "custom_name.docx")
  res <- convertxlsform(src, output = out, verbose = FALSE)
  expect_equal(basename(res), "custom_name.docx")
  expect_true(file.exists(out))
})

test_that("a missing output folder is an error", {
  stub_render_docx()
  src <- copy_example()
  expect_error(
    convertxlsform(src, output = file.path(tempdir(), "no_such_dir", "x.docx"),
                   verbose = FALSE),
    "does not exist"
  )
})

test_that("format is inferred from the output extension", {
  skip_if_not(has_libreoffice())
  skip_on_cran()
  stub_render_docx()
  src <- copy_example()
  out <- file.path(withr::local_tempdir(), "x.pdf")
  res <- convertxlsform(src, output = out, verbose = FALSE)
  expect_match(res, "x\\.pdf$")
  expect_true(file.exists(out))
})

test_that("a mismatch between output and format is an error", {
  stub_render_docx()
  src <- copy_example()
  expect_error(
    convertxlsform(src, output = file.path(tempdir(), "x.docx"),
                   format = "pdf", verbose = FALSE),
    "disagree"
  )
  expect_error(
    convertxlsform(src, output = file.path(tempdir(), "x.txt"), verbose = FALSE),
    "must end in"
  )
  expect_error(convertxlsform(src, format = "odt", verbose = FALSE),
               "must be")
})

test_that("existing files are protected unless overwrite = TRUE", {
  stub_render_docx()
  src <- copy_example()
  first <- convertxlsform(src, verbose = FALSE)
  err <- expect_error(convertxlsform(src, verbose = FALSE), "already exists")
  expect_match(conditionMessage(err), "myform.docx")
  expect_match(conditionMessage(err), "overwrite = TRUE")
  expect_equal(convertxlsform(src, overwrite = TRUE, verbose = FALSE), first)
})

test_that("both formats share a base name", {
  skip_if_not(has_libreoffice())
  skip_on_cran()
  stub_render_docx()
  src <- copy_example()
  res <- convertxlsform(src, format = c("docx", "pdf"), verbose = FALSE)
  expect_equal(basename(res), c("myform.docx", "myform.pdf"))
  expect_true(all(file.exists(res)))
  res2 <- convertxlsform(src, format = "both", overwrite = TRUE, verbose = FALSE)
  expect_equal(res2, res)
})

test_that("pdf-only output leaves no docx behind", {
  skip_if_not(has_libreoffice())
  skip_on_cran()
  stub_render_docx()
  src <- copy_example()
  res <- convertxlsform(src, format = "pdf", verbose = FALSE)
  expect_equal(basename(res), "myform.pdf")
  expect_false(file.exists(sub("pdf$", "docx", res)))
})

test_that("a missing input names the path and reports convertxlsform()", {
  stub_render_docx()
  err <- expect_error(convertxlsform("nope/missing.xlsx", verbose = FALSE),
                      "not found")
  expect_match(conditionMessage(err), "missing.xlsx")
  expect_match(conditionMessage(err), "Working directory")
  expect_equal(rlang::call_name(err$call), "convertxlsform")
})

test_that("option errors report convertxlsform() too", {
  stub_render_docx()
  src <- copy_example()
  err <- expect_error(convertxlsform(src, page_size = "B5", verbose = FALSE))
  expect_equal(rlang::call_name(err$call), "convertxlsform")
})

test_that("verbose = FALSE is silent and the path is returned invisibly", {
  stub_render_docx()
  src <- copy_example()
  expect_silent(convertxlsform(src, verbose = FALSE))
  withr::defer(unlink(sub("xlsx$", "docx", src)))
  vis <- withVisible(convertxlsform(src, overwrite = TRUE, verbose = FALSE))
  expect_false(vis$visible)
  expect_type(vis$value, "character")
})

test_that("verbose = TRUE reports the steps and the timing", {
  stub_render_docx()
  src <- copy_example()
  msgs <- testthat::capture_messages(convertxlsform(src))
  msgs <- paste(msgs, collapse = "\n")
  expect_match(msgs, "Reading")
  expect_match(msgs, "Language: .*; [0-9]+ questions")
  expect_match(msgs, "Wrote .*myform\\.docx.* in [0-9.]+ s")
})

test_that("the real renderer produces a readable document", {
  fn <- get(".render_docx", envir = asNamespace("convertxlsform"))
  skip_if(any(grepl("stub", deparse(fn), fixed = TRUE)),
          "renderer not available yet")
  src <- copy_example()
  res <- convertxlsform(src, verbose = FALSE)
  doc <- officer::read_docx(res)
  expect_gt(nrow(officer::docx_summary(doc)), 5)
})
