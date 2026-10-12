# Tests for the Word renderer (R/render_docx.R, R/answer_layouts.R).

example_model <- function(example, lang = NULL, ...) {
  p <- system.file("extdata", example, paste0(example, ".xlsx"),
                   package = "convertxlsform")
  skip_if(!nzchar(p), "example not installed")
  form <- read_quiet(p, lang = lang, media_root = dirname(p))
  opts <- test_opts(media_root = dirname(p), lang = lang, ...)
  list(model = .build_model(form, opts), opts = opts)
}

render_to <- function(m, ...) {
  path <- tempfile(fileext = ".docx")
  opts <- m$opts
  if (...length() > 0L) {
    opts <- utils::modifyList(unclass(opts), list(...))
  }
  .render_docx(m$model, opts, path)
  path
}

docx_text <- function(path) {
  s <- officer::docx_summary(officer::read_docx(path))
  s$text[!is.na(s$text)]
}

zip_part <- function(path, pattern) {
  dir <- tempfile()
  dir.create(dir)
  files <- utils::unzip(path, exdir = dir)
  f <- files[grepl(pattern, files)]
  vapply(f, function(x) paste(readLines(x, encoding = "UTF-8", warn = FALSE),
                              collapse = "\n"), "")
}

test_that(".text_runs handles markdown, html and line breaks", {
  fp <- officer::fp_text_lite(font.size = 10)
  r <- .text_runs("plain **bold** and *it* and _it2_", fp)
  txt <- vapply(r, function(x) x$value, "")
  expect_equal(paste(txt, collapse = ""), "plain bold and it and it2")
  bold <- vapply(r, function(x) isTRUE(x$pr$bold), NA)
  expect_equal(txt[bold], "bold")

  r <- .text_runs("a <span style=\"color:red\">b</span>\nc", fp)
  expect_length(r, 3L)
  expect_equal(r[[1]]$value, "a b")
  expect_s3_class(r[[2]], "run_linebreak")
  expect_equal(r[[3]]$value, "c")

  # underscores inside names and blanks are literal
  r <- .text_runs("[hh_size] ____", fp)
  expect_equal(paste(vapply(r, function(x) x$value, ""), collapse = ""),
               "[hh_size] ____")
  expect_length(.text_runs(NA_character_, fp), 0L)
})

test_that("all_types renders with title, serials in order and choices", {
  m <- example_model("all_types", "en")
  path <- render_to(m)
  expect_true(file.exists(path))
  txt <- docx_text(path)
  expect_true(m$model$meta$title %in% txt)

  secq <- Filter(function(b) b$block %in% c("section", "question") &&
                   !is.na(b$serial), m$model$blocks)
  expected <- vapply(secq, function(b) {
    paste0(b$serial, ". ", substr(strsplit(b$label, "\n")[[1]][1], 1, 12))
  }, "")
  pos <- vapply(expected, function(e) which(startsWith(txt, e))[1], 1L)
  expect_false(anyNA(pos))
  expect_true(all(diff(pos) > 0))

  # choice labels
  for (lab in c("Tap", "Mobile phone", "Own farm", "Village C2")) {
    expect_true(any(startsWith(trimws(txt), lab)), info = lab)
  }
  # required asterisk runs and legend
  expect_true(any(grepl(" \\*$", txt)))
  expect_true("* Required question" %in% txt)
})

test_that("Ask if lines follow include_relevants", {
  on <- render_to(example_model("all_types", "en"))
  expect_true(any(grepl("^Ask if: ", docx_text(on))))
  off <- render_to(example_model("all_types", "en", include_relevants = FALSE))
  expect_false(any(grepl("^Ask if: ", docx_text(off))))
})

test_that("French strings appear in the French render", {
  m <- example_model("all_types", "fr")
  txt <- docx_text(render_to(m))
  expect_true(any(grepl("^Poser si", txt)))
  expect_true(any(grepl("^R\u00e9p\u00e9ter pour chaque", txt)))
  expect_true(any(grepl("^Occurrence n", txt)))
  foot <- zip_part(render_to(m), "footer")
  expect_match(foot, "sur", all = FALSE)
})

test_that("repeat_copies duplicates repeat content", {
  m <- example_model("all_types", "en")
  one <- docx_text(render_to(m))
  two <- docx_text(render_to(m, repeat_copies = 2L))
  nm <- "4.1. First name of member"
  expect_equal(sum(startsWith(one, nm)), 1L)
  expect_equal(sum(startsWith(two, nm)), 2L)
  expect_equal(sum(startsWith(one, "Instance no.")), 1L)
  expect_equal(sum(startsWith(two, "Instance no.")), 2L)
  # the heading and the banner are not duplicated
  expect_equal(sum(startsWith(two, "Repeat for each")), 1L)
})

test_that("images are embedded", {
  m <- example_model("all_types", "en")
  path <- render_to(m)
  files <- utils::unzip(path, list = TRUE)$Name
  media <- files[grepl("^word/media/", files)]
  expect_gte(length(media), 4L)
  # a missing image is a note, not an error
  expect_true(any(grepl("^\\[image: ", docx_text(path))) ||
                length(media) >= 4L)
})

test_that("header and footer carry title, form id and page fields", {
  m <- example_model("all_types", "en")
  path <- render_to(m)
  hdr <- zip_part(path, "word/header[0-9]*\\.xml$")
  ftr <- zip_part(path, "word/footer[0-9]*\\.xml$")
  expect_match(hdr, m$model$meta$title, fixed = TRUE, all = FALSE)
  expect_match(ftr, "PAGE", all = FALSE)
  expect_match(ftr, "NUMPAGES", all = FALSE)
  expect_match(ftr, m$model$meta$form_id, fixed = TRUE, all = FALSE)
})

test_that("multilang renders in Bengali with the complex-script font", {
  m <- example_model("multilang", "bn")
  path <- render_to(m)
  txt <- docx_text(path)
  sec <- Filter(function(b) b$block == "question", m$model$blocks)[[1]]
  expect_true(any(grepl(substr(sec$label, 1, 6), txt, fixed = TRUE)))
  doc <- zip_part(path, "word/document.xml")
  expect_match(doc, "w:cs=\"Noto Sans Bengali\"", fixed = TRUE)
})

test_that("a user template without cxf styles falls back to Normal", {
  tpl <- tempfile(fileext = ".docx")
  print(officer::read_docx(), target = tpl)
  m <- example_model("all_types", "en")
  path <- tempfile(fileext = ".docx")
  expect_no_error(.render_docx(m$model, utils::modifyList(
    unclass(m$opts), list(template = tpl)), path))
  txt <- docx_text(path)
  expect_true(any(startsWith(txt, "1. Does the respondent")))
  s <- officer::docx_summary(officer::read_docx(path))
  expect_false(any(s$style_name %in% c("cxf Question", "cxf Hint"),
                   na.rm = TRUE))
})

test_that("landscape letter page setup is applied", {
  m <- example_model("all_types", "en", page_size = "letter",
                     orientation = "landscape", margins_in = 0.75)
  doc <- zip_part(render_to(m), "word/document.xml")
  expect_match(doc, "w:orient=\"landscape\"", fixed = TRUE)
  expect_match(doc, "w:w=\"15840\"", fixed = TRUE)
})

test_that("small forms render with codes hidden and variable names shown", {
  f <- small_form()
  opts <- test_opts(include_choice_name = FALSE, show_variable_names = TRUE)
  model <- .build_model(f, opts)
  path <- tempfile(fileext = ".docx")
  .render_docx(model, opts, path)
  txt <- docx_text(path)
  expect_true(any(grepl("\\[consent\\]", txt)))
  expect_true("Radio" %in% txt)
  expect_false("radio" %in% txt)
})
