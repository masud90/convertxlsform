# Tests for the Word renderer (R/render_docx.R, R/answer_layouts.R).

example_model <- function(example, lang = NULL, ...) {
  p <- system.file("extdata", example, paste0(example, ".xlsx"),
                   package = "convertxlsform")
  skip_if(!nzchar(p), "example not installed")
  form <- read_quiet(p, lang = lang, media_root = dirname(p))
  opts <- test_opts(media_root = dirname(p), lang = lang, ...)
  list(model = .build_model(form, opts), opts = opts)
}

example_model_from <- function(path, ...) {
  form <- read_quiet(path)
  opts <- test_opts(...)
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
  # officer 0.7 lists table cells after the paragraphs; restore document order
  s <- s[order(s$doc_index, s$row_id, s$cell_id), ]
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

test_that("the minimal form renders with title, serials and choices", {
  m <- example_model("minimal")
  path <- render_to(m)
  expect_true(file.exists(path))
  txt <- docx_text(path)
  expect_true(m$model$meta$title %in% txt)
  expect_true(any(startsWith(txt, "1. What is the respondent")))
  expect_true(any(startsWith(txt, "5.1. Where is the tap")))
  expect_true(any(startsWith(trimws(txt), "Tap")))
  expect_true(any(grepl("^Ask if: ", txt)))
  expect_equal(sum(grepl("Enter an age between 0 and 120", txt)), 1L)
})

test_that("all_types renders with title, serials in order and choices", {
  skip_on_cran()
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
  skip_on_cran()
  on <- render_to(example_model("all_types", "en"))
  expect_true(any(grepl("^Ask if: ", docx_text(on))))
  off <- render_to(example_model("all_types", "en", include_relevants = FALSE))
  expect_false(any(grepl("^Ask if: ", docx_text(off))))
})

test_that("French strings appear in the French render", {
  skip_on_cran()
  m <- example_model("all_types", "fr", repeat_layout = "copies")
  txt <- docx_text(render_to(m))
  expect_true(any(grepl("^Poser si", txt)))
  expect_true(any(grepl("^R\u00e9p\u00e9ter pour chaque", txt)))
  expect_true(any(grepl("^Occurrence n", txt)))
  foot <- zip_part(render_to(m), "footer")
  expect_match(foot, "sur", all = FALSE)
})

test_that("repeat_copies duplicates repeat content", {
  skip_on_cran()
  m <- example_model("all_types", "en", repeat_layout = "copies")
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
  skip_on_cran()
  m <- example_model("all_types", "en")
  path <- render_to(m)
  files <- utils::unzip(path, list = TRUE)$Name
  media <- files[grepl("^word/media/", files)]
  expect_gte(length(media), 4L)
  # a missing image is a note, not an error
  expect_true(any(grepl("^\\[image: ", docx_text(path))) ||
                length(media) >= 4L)
})

test_that("choice rows with an image are top-aligned", {
  skip_on_cran()
  m <- example_model("all_types", "en")
  doc <- zip_part(render_to(m), "word/document\\.xml$")
  rows <- regmatches(doc, gregexpr("(?s)<w:tr[ >].*?</w:tr>", doc, perl = TRUE))[[1]]
  with_img <- rows[grepl("<w:drawing", rows, fixed = TRUE)]
  skip_if(length(with_img) == 0L, "no image rows rendered")
  expect_false(any(grepl('<w:vAlign w:val="center"', with_img, fixed = TRUE)))
  expect_true(any(grepl('<w:vAlign w:val="top"', with_img, fixed = TRUE)))
})

test_that("header and footer carry title, form id and page fields", {
  m <- example_model("minimal")
  path <- render_to(m)
  hdr <- zip_part(path, "word/header[0-9]*\\.xml$")
  ftr <- zip_part(path, "word/footer[0-9]*\\.xml$")
  expect_match(hdr, m$model$meta$title, fixed = TRUE, all = FALSE)
  expect_match(ftr, "PAGE", all = FALSE)
  expect_match(ftr, "NUMPAGES", all = FALSE)
  expect_match(ftr, m$model$meta$form_id, fixed = TRUE, all = FALSE)
})

test_that("multilang renders in Bengali with the complex-script font", {
  skip_on_cran()
  m <- example_model("multilang", "bn")
  path <- render_to(m)
  txt <- docx_text(path)
  sec <- Filter(function(b) b$block == "question", m$model$blocks)[[1]]
  expect_true(any(grepl(substr(sec$label, 1, 6), txt, fixed = TRUE)))
  doc <- zip_part(path, "word/document.xml")
  # the first installed Bengali font, e.g. Nirmala UI on Windows
  cs <- m$model$meta$font_family_cs
  expect_true(cs %in% .cs_scripts()$bengali$fonts)
  expect_match(doc, paste0("w:cs=\"", cs, "\""), fixed = TRUE)
})

test_that("a user template without cxf styles falls back to Normal", {
  tpl <- tempfile(fileext = ".docx")
  print(officer::read_docx(), target = tpl)
  m <- example_model("minimal")
  path <- tempfile(fileext = ".docx")
  expect_no_error(.render_docx(m$model, utils::modifyList(
    unclass(m$opts), list(template = tpl)), path))
  txt <- docx_text(path)
  expect_true(any(startsWith(txt, "1. What is the respondent")))
  s <- officer::docx_summary(officer::read_docx(path))
  expect_false(any(s$style_name %in% c("cxf Question", "cxf Hint"),
                   na.rm = TRUE))
})

test_that("landscape letter page setup is applied", {
  m <- example_model("minimal", page_size = "letter",
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

test_that("every Bengali run, including code-table cells, uses the cs font", {
  skip_on_cran()
  m <- example_model("multilang", "bn")
  doc <- zip_part(render_to(m), "word/document.xml")
  runs <- regmatches(doc, gregexpr("<w:r>.*?</w:r>", doc))[[1L]]
  bn <- runs[grepl("[\u0980-\u09ff]", runs)]
  expect_true(length(bn) > 0L)
  expect_false(any(grepl("w:cs=\"Arial\"", bn, fixed = TRUE)))
})

test_that("circle symbols are drawn larger than square symbols", {
  expect_equal(.symbol_size(.cxf_symbols()$square, 11), 14)
  expect_equal(.symbol_size(.cxf_symbols()$circle, 11), 18)
  expect_equal(.cxf_choice_image_max_in, 0.6)
})

test_that("a relative output path still gets the bumped complex-script size", {
  skip_on_cran()
  skip_if_not_installed("withr")
  m <- example_model("multilang", "bn")
  abs_path <- render_to(m)
  d <- tempfile("cxfwd")
  dir.create(d)
  withr::local_dir(d)
  .render_docx(m$model, m$opts, "rel_out.docx")
  expect_true(file.exists(file.path(d, "rel_out.docx")))
  szcs <- function(p) {
    regmatches(zip_part(p, "word/document.xml"),
               gregexpr("w:szCs w:val=\"[0-9]+\"",
                        zip_part(p, "word/document.xml")))[[1L]]
  }
  expect_equal(szcs("rel_out.docx"), szcs(abs_path))
  expect_equal(normalizePath(getwd(), winslash = "/"),
               normalizePath(d, winslash = "/"))
})

test_that("markdown links and heading markers are cleaned", {
  s <- .text_segments("See [the guide](https://x.org/a_b) now")
  expect_equal(paste(s$text, collapse = ""), "See the guide now")
  s <- .text_segments("## Section title\nplain *it*\n# Top")
  expect_equal(s$text[!s$br], c("Section title", "plain ", "it", "Top"))
  expect_equal(s$bold[!s$br], c(TRUE, FALSE, FALSE, TRUE))
  # a hash that is not a heading marker stays
  expect_equal(.text_segments("Item #5")$text, "Item #5")
})

test_that("a very narrow text width does not give negative column widths", {
  m <- example_model("minimal")
  path <- expect_no_error(render_to(m, text_width_in = 2.5))
  expect_true(file.exists(path))
  doc <- zip_part(path, "word/document.xml")
  expect_false(grepl("w:w=\"-", doc, fixed = TRUE))
})

# ---- 0.4.0 layouts ------------------------------------------------------------

table_form <- function() {
  skip_if_no_writexl()
  sv <- fdf(
    type = c("integer", "begin_repeat", "text", "select_one sex", "date",
             "end_repeat", "begin_group", "select_one agree",
             "select_one agree", "end_group", "select_one agree",
             "select_one agree", "select_one agree"),
    name = c("n", "mem", "mname", "sex", "dob", NA, "att", "a1", "a2", NA,
             "hdr", "b1", "b2"),
    label = c("How many?", "Members", "Name", "Sex", "Date of birth", NA,
              "Attitudes", "Schools are good", "Clinics are good", NA,
              "Rate", "Water is clean", "Power is stable"),
    appearance = c(NA, NA, NA, NA, NA, NA, "table-list", NA, NA, NA,
                   "label", "list-nolabel", "list-nolabel"),
    relevant = c(NA, NA, NA, NA, "${sex} = '1'", NA, NA, NA, NA, NA, NA, NA,
                 NA),
    repeat_count = c(NA, "3", rep(NA, 11))
  )
  ch <- fdf(list_name = c("sex", "sex", "agree", "agree", "agree"),
            name = c("1", "2", "1", "2", "3"),
            label = c("Male", "Female", "Agree", "Neutral", "Disagree"))
  make_form(sv, ch, fdf(form_title = "Tables"))
}

test_that("the label row of a grid takes no number", {
  out <- xlsform_outline(table_form())
  expect_true(is.na(out$serial[out$name == "hdr"]))
  expect_equal(out$serial[out$name %in% c("b1", "b2")], c("4", "5"))
})

test_that("table-list groups and label rows render as grids", {
  skip_on_cran()
  p <- table_form()
  form <- read_quiet(p)
  opts <- test_opts()
  m <- .build_model(form, opts)
  g <- .grid_from_section(m$blocks, which(block_types(m) == "section")[2L])
  expect_length(g$rows, 2L)
  hdr <- which(vapply(m$blocks, function(b) identical(b$name, "hdr"), NA))
  g2 <- .grid_from_label(m$blocks, hdr)
  expect_equal(vapply(g2$rows, function(b) b$name, ""), c("b1", "b2"))
  s <- officer::docx_summary(officer::read_docx(render_to(list(model = m,
                                                               opts = opts))))
  cells <- s$text[s$content_type == "table cell"]
  expect_true(all(c("Agree", "Neutral", "Disagree") %in%
                    sub("\\s*[0-9]$", "", cells)))
  expect_true(any(grepl("^3\\.1\\. Schools are good", cells)))
  expect_true(any(grepl("^4\\. Water is clean", cells)))
  # the questions of a grid are not printed again as paragraphs
  paras <- s$text[s$content_type == "paragraph"]
  expect_false(any(grepl("Schools are good", paras)))
  expect_false(any(grepl("^Rate", paras)))
})

test_that("a grid needs the same choices on every row", {
  b <- function(name, labels) {
    list(block = "question", kind = "select_one", name = name,
         answer = list(source_found = TRUE, choices = data.frame(
           name = seq_along(labels), label = labels, write_in = FALSE,
           image = I(rep(list(NULL), length(labels))))))
  }
  expect_false(is.na(.grid_key(b("a", c("Yes", "No")))))
  expect_false(identical(.grid_key(b("a", c("Yes", "No"))),
                         .grid_key(b("b", c("Yes", "Maybe")))))
  expect_true(is.na(.grid_key(b("a", as.character(1:9)))))
})

test_that("repeats render as rosters in auto mode", {
  skip_on_cran()
  p <- table_form()
  m <- example_model_from(p)
  s <- officer::docx_summary(officer::read_docx(render_to(m)))
  cells <- s$text[s$content_type == "table cell"]
  expect_true(any(grepl("^2\\.1\\. Name", cells)))
  expect_true(any(grepl("dd/mm/yyyy", cells)))
  # repeat_count 3 gives three numbered rows
  expect_true(all(c("1", "2", "3") %in% cells))
  expect_false("4" %in% cells[!grepl("\\.", cells)])
  paras <- s$text[s$content_type == "paragraph"]
  expect_true(any(grepl("^Q2\\.2: 1 = Male; 2 = Female", paras)))
  expect_true(any(grepl("^Q2\\.3 Ask if:", paras)))
  expect_false(any(grepl("^Instance no\\.", paras)))
  copies <- officer::docx_summary(officer::read_docx(
    render_to(m, repeat_layout = "copies")))
  expect_true(any(grepl("^Instance no\\.", copies$text)))
})

test_that(".use_roster follows repeat_layout and the column limits", {
  q <- list(block = "question", kind = "text", image = NULL)
  note <- list(block = "note")
  o <- function(x) list(repeat_layout = x)
  expect_true(.use_roster(rep(list(q), 6), o("auto")))
  expect_false(.use_roster(rep(list(q), 7), o("auto")))
  expect_true(.use_roster(rep(list(q), 7), o("roster")))
  expect_false(.use_roster(rep(list(q), 2), o("copies")))
  expect_false(.use_roster(list(q, note), o("roster")))
  expect_false(.use_roster(list(), o("roster")))
})

test_that("right-to-left forms get bidi paragraphs, tables and runs", {
  skip_on_cran()
  skip_if_no_writexl()
  sv <- fdf(type = c("select_one yn", "text"), name = c("q1", "q2"),
            "label::Arabic (ar)" = c("هل توافق؟",
                                     "الاسم"))
  ch <- fdf(list_name = c("yn", "yn"), name = c("1", "2"),
            "label::Arabic (ar)" = c("نعم", "لا"))
  p <- make_form(sv, ch)
  m <- example_model_from(p)
  expect_true(m$model$meta$rtl)
  doc <- zip_part(render_to(m), "word/document\\.xml$")
  expect_match(doc, "<w:bidi/>")
  expect_match(doc, "<w:bidiVisual/>")
  expect_match(doc, "<w:rtl/>")
  off <- zip_part(render_to(example_model_from(p, rtl = FALSE)),
                  "word/document\\.xml$")
  expect_false(grepl("<w:bidi/>", off))
  # the document still opens
  expect_s3_class(officer::read_docx(render_to(m)), "rdocx")
})

test_that(".xml_insert_child keeps schema order and is idempotent", {
  x <- "<w:pPr><w:pStyle w:val=\"a\"/><w:jc w:val=\"left\"/></w:pPr><w:pPr/>"
  y <- .xml_insert_child(x, "w:pPr", "<w:bidi/>", c("w:jc"))
  expect_equal(y, paste0("<w:pPr><w:pStyle w:val=\"a\"/><w:bidi/>",
                         "<w:jc w:val=\"left\"/></w:pPr>",
                         "<w:pPr><w:bidi/></w:pPr>"))
  expect_equal(.xml_insert_child(y, "w:pPr", "<w:bidi/>", c("w:jc")), y)
  r <- .rtl_xml("<w:p><w:r><w:t>م</w:t></w:r><w:r><w:t>a</w:t></w:r></w:p>")
  expect_equal(lengths(regmatches(r, gregexpr("<w:rtl/>", r))), 1L)
  expect_match(r, "^<w:p><w:pPr><w:bidi/></w:pPr>")
})

test_that("entity lists appear in the front matter", {
  skip_on_cran()
  skip_if_no_writexl()
  sv <- fdf(type = "text", name = "hh", label = "Head", save_to = "hh_name")
  p <- make_form(sv, settings = fdf(form_title = "E"),
                 entities = fdf(list_name = "households", label = "${hh}"))
  m <- example_model_from(p, show_variable_names = TRUE)
  expect_equal(m$model$meta$entities, "households")
  txt <- docx_text(render_to(m))
  expect_true(any(grepl("Entity list: households", txt, fixed = TRUE)))
  expect_true(any(grepl("[hh → hh_name]", txt, fixed = TRUE)))
})
