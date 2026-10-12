make_docx <- function(path = tempfile(fileext = ".docx")) {
  officer::read_docx() %>%
    officer::body_add_par("Hello") %>%
    print(target = path)
  path
}

test_that(".find_soffice() honours the option", {
  dir <- withr::local_tempdir()
  fake <- file.path(dir, "fake-soffice")
  file.create(fake)
  withr::local_options(convertxlsform.soffice = fake)
  expect_equal(basename(.find_soffice()), "fake-soffice")
  expect_true(has_libreoffice())
})

test_that(".find_soffice() honours the environment variable", {
  dir <- withr::local_tempdir()
  fake <- file.path(dir, "env-soffice")
  file.create(fake)
  withr::local_options(convertxlsform.soffice = NULL)
  withr::local_envvar(CONVERTXLSFORM_SOFFICE = fake)
  expect_equal(basename(.find_soffice()), "env-soffice")
})

test_that(".find_soffice() ignores an option pointing to a missing file", {
  withr::local_options(convertxlsform.soffice = "/no/such/soffice")
  withr::local_envvar(CONVERTXLSFORM_SOFFICE = "")
  res <- .find_soffice()
  expect_true(is.na(res) || !identical(res, "/no/such/soffice"))
})

test_that("docx_to_pdf() validates its inputs", {
  dir <- withr::local_tempdir()
  expect_error(docx_to_pdf(file.path(dir, "missing.docx")), "does not exist")

  txt <- file.path(dir, "a.txt")
  writeLines("x", txt)
  expect_error(docx_to_pdf(txt), "docx")

  docx <- make_docx(file.path(dir, "a.docx"))
  existing <- file.path(dir, "a.pdf")
  writeLines("x", existing)
  expect_error(docx_to_pdf(docx, existing), "already exists")
  expect_equal(readLines(existing), "x")
})

test_that("a failed LibreOffice run aborts with its output", {
  dir <- withr::local_tempdir()
  fake <- file.path(dir, "fake-soffice")
  file.create(fake)
  Sys.chmod(fake, "755")
  withr::local_options(convertxlsform.soffice = fake)
  docx <- make_docx(file.path(dir, "a.docx"))
  expect_error(
    docx_to_pdf(docx, file.path(dir, "out.pdf"), verbose = FALSE),
    "did not produce"
  )
})

test_that("a first LibreOffice start that makes no PDF is retried once", {
  skip_on_os("windows")
  dir <- withr::local_tempdir()
  counter <- file.path(dir, "counter")
  fake <- file.path(dir, "fake-soffice")
  writeLines(c(
    "#!/bin/sh",
    sprintf("COUNTER='%s'", counter),
    "echo x >> \"$COUNTER\"",
    "n=$(wc -l < \"$COUNTER\" | tr -d ' ')",
    "if [ \"$n\" -lt 2 ]; then exit 1; fi",
    "outdir=''; prev=''",
    "for a in \"$@\"; do",
    "  if [ \"$prev\" = '--outdir' ]; then outdir=\"$a\"; fi",
    "  prev=\"$a\"",
    "  last=\"$a\"",
    "done",
    "base=$(basename \"$last\" .docx)",
    "printf '%%PDF-1.4 fake' > \"$outdir/$base.pdf\""
  ), fake)
  Sys.chmod(fake, "755")
  withr::local_options(convertxlsform.soffice = fake)
  docx <- make_docx(file.path(dir, "a.docx"))
  target <- file.path(dir, "out.pdf")
  expect_no_error(docx_to_pdf(docx, target, verbose = FALSE))
  expect_true(file.exists(target))
  expect_equal(length(readLines(counter)), 2L)
})

test_that("docx_to_pdf() converts with LibreOffice", {
  skip_on_cran()
  skip_if_not(has_libreoffice())
  dir <- withr::local_tempdir()
  docx <- make_docx(file.path(dir, "a.docx"))
  target <- file.path(dir, "out.pdf")

  res <- docx_to_pdf(docx, target, engine = "libreoffice", verbose = FALSE)
  expect_equal(res, target)
  expect_true(file.exists(target))
  expect_equal(rawToChar(readBin(target, "raw", 4L)), "%PDF")

  expect_error(docx_to_pdf(docx, target, verbose = FALSE), "already exists")
  expect_no_error(docx_to_pdf(docx, target, overwrite = TRUE, verbose = FALSE))

  default <- docx_to_pdf(docx, verbose = FALSE)
  expect_equal(basename(default), "a.pdf")
  expect_equal(normalizePath(dirname(default)), normalizePath(dir))
  expect_true(file.exists(default))
})

test_that("engine = 'word' explains the missing doconv package", {
  skip_if(requireNamespace("doconv", quietly = TRUE))
  docx <- make_docx()
  expect_error(
    docx_to_pdf(docx, tempfile(fileext = ".pdf"), engine = "word"),
    "doconv"
  )
})

test_that(".soffice_ld_path puts the LibreOffice program folder first", {
  d <- withr::local_tempdir()
  prog <- file.path(d, "program")
  dir.create(prog)
  file.create(file.path(prog, "soffice.bin"))
  bin <- file.path(prog, "soffice")
  file.create(bin)
  real_prog <- dirname(normalizePath(bin, mustWork = FALSE))
  expect_equal(.soffice_ld_path(bin, NA), real_prog)
  expect_equal(.soffice_ld_path(bin, ""), real_prog)
  expect_equal(
    .soffice_ld_path(bin, "/opt/R/lib"),
    paste(real_prog, "/opt/R/lib", sep = ":")
  )
})
