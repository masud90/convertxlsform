write_csv_dir <- function(...) {
  d <- tempfile("csvdir")
  dir.create(d)
  files <- list(...)
  for (nm in names(files)) {
    writeLines(files[[nm]], file.path(d, nm), useBytes = TRUE)
  }
  d
}

test_that("choice lists come from the choices sheet in order", {
  skip_if_no_writexl()
  f <- small_form()
  a <- .choice_list(f, "assets")
  expect_equal(a$name, c("radio", "tv", "bike"))
  expect_equal(a$label, c("Radio", "TV", "Bike"))
  expect_named(a, c("name", "label", "image", "row"))
  expect_equal(nrow(.choice_list(f, "yesno")), 2L)
  expect_equal(nrow(.choice_list(f, "nope")), 0L)
  expect_equal(nrow(.choice_list(f, NA_character_)), 0L)
})

test_that("a blank choice label falls back to the name", {
  skip_if_no_writexl()
  survey <- fdf(type = "select_one l", name = "q", label = "Q")
  choices <- fdf(list_name = c("l", "l"), name = c("a", "b"),
                 label = c("Alpha", NA))
  f <- read_quiet(make_form(survey, choices))
  expect_equal(.choice_list(f, "l")$label, c("Alpha", "b"))
})

test_that("CSV choices are read with the chosen language", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  en <- read_quiet(p, lang = "en")
  fr <- read_quiet(p, lang = "fr")
  a <- .file_choices(en, "villages.csv")
  expect_true(a$found)
  expect_true(is.na(a$note))
  expect_equal(nrow(a$choices), 5L)
  expect_equal(a$choices$name[1], "village_1")
  expect_equal(a$choices$label[1], "Village 1")
  expect_named(a$choices, c("name", "label", "image"))
  expect_equal(basename(a$path), "villages.csv")
  # the label column follows the language (same labels in this file, so use
  # a purpose-built file below for the contrast)
  b <- .file_choices(fr, "villages.csv")
  expect_equal(nrow(b$choices), 5L)
})

test_that("XML choices are read from the example file", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  skip_if(!file.exists(file.path(dirname(p), "facilities.xml")))
  a <- .file_choices(read_quiet(p, lang = "fr"), "facilities.xml")
  expect_true(a$found)
  expect_true(is.na(a$note))
  expect_named(a$choices, c("name", "label", "image"))
  expect_equal(a$choices$name,
               c("school", "clinic", "market", "water_point"))
  expect_equal(a$choices$label[2], "Health centre & pharmacy")
  expect_true(all(is.na(a$choices$image)))
})

test_that("XML choice files: entities, CDATA, comments and blanks", {
  skip_if_no_writexl()
  d <- write_csv_dir(
    "odk.xml" = c(
      "\ufeff<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
      "<!-- a comment with <item><name>x</name></item> inside -->",
      "<root>",
      "  <item><name>a</name><label>A &lt;1&gt; &quot;q&quot;</label></item>",
      "  <item>",
      "    <name>b</name>",
      "    <label><![CDATA[B & <b>]]></label>",
      "    <extra/>",
      "  </item>",
      "  <item><name>c</name><label>Caf&#233; &#x41;</label></item>",
      "  <item><name>d</name><label/></item>",
      "  <item><label>no name</label></item>",
      "</root>"
    ),
    "multi.xml" = c(
      "<root>",
      "<item><name>1</name><label>Cat</label><district>n</district></item>",
      "<item><name>2</name><label></label></item>",
      "</root>"
    ),
    "nolabel.xml" = "<root><item><id>7</id></item></root>"
  )
  f <- read_quiet(make_form(fdf(type = "text", name = "q", label = "Q")),
                  media_root = d)
  a <- .file_choices(f, "odk.xml")
  expect_true(a$found)
  expect_equal(a$choices$name, c("a", "b", "c", "d"))
  expect_equal(a$choices$label,
               c("A <1> \"q\"", "B & <b>", "Caf\u00e9 A", "d"))
  m <- .file_choices(f, "multi.xml")
  expect_equal(m$choices$label, c("Cat", "2"))
  # without a name element the first column is the value
  n <- .file_choices(f, "nolabel.xml")
  expect_equal(n$choices$name, "7")
  expect_equal(n$choices$label, "7")
})

test_that("a select_one_from_file with an XML file lists its choices", {
  skip_if_no_writexl()
  d <- write_csv_dir("v.xml" = c(
    "<root>",
    "<item><name>v1</name><label>Village 1</label></item>",
    "<item><name>v2</name><label>Village 2</label></item>",
    "</root>"
  ))
  survey <- fdf(type = c("select_one_from_file v.xml", "text"),
                name = c("v", "t"), label = c("Village", "T"),
                relevant = c(NA, "${v} = 'v2'"))
  f <- read_quiet(make_form(survey), media_root = d)
  sc <- .select_choices(f, 1L)
  expect_true(sc$source_found)
  expect_true(is.na(sc$note))
  expect_equal(sc$choices$label, c("Village 1", "Village 2"))
  # skip logic shows the label from the file
  expect_equal(.humanize_expr("${v} = 'v2'", f, "relevant"),
               "Q1 (v) is \"Village 2\"")
})

test_that("the label column follows the language and falls back", {
  skip_if_no_writexl()
  d <- write_csv_dir(
    "multi.csv" = c("name,label::English (en),label::French (fr)",
                    "1,Cat,Chat", "2,Dog,Chien", "3,Bird,"),
    "plain.csv" = c("name,label", "1,One", "2,Two"),
    "nolabel.csv" = c("id,code", "1,x"),
    "oneling.csv" = c("name,label::French (fr)", "1,Un")
  )
  survey <- fdf(type = c("select_one_from_file multi.csv",
                         "select_one_from_file plain.csv"),
                name = c("a", "b"), `label::English (en)` = c("A", "B"),
                `label::French (fr)` = c("A fr", "B fr"))
  path <- make_form(survey, settings = fdf(default_language = "English (en)"))
  en <- read_quiet(path, media_root = d, lang = "en")
  fr <- read_quiet(path, media_root = d, lang = "fr")
  expect_equal(.file_choices(en, "multi.csv")$choices$label,
               c("Cat", "Dog", "Bird"))
  # French is empty for the third row: falls back to the default language
  expect_equal(.file_choices(fr, "multi.csv")$choices$label,
               c("Chat", "Chien", "Bird"))
  # a plain label column serves every language
  expect_equal(.file_choices(fr, "plain.csv")$choices$label, c("One", "Two"))
  # no label column at all: the first column is the name, label = name
  nl <- .file_choices(en, "nolabel.csv")
  expect_equal(nl$choices$name, "1")
  expect_equal(nl$choices$label, "1")
  # only another language available: use it
  expect_equal(.file_choices(en, "oneling.csv")$choices$label, "Un")
})

test_that("a byte order mark and UTF-8 text are handled", {
  skip_if_no_writexl()
  d <- tempfile("csvdir")
  dir.create(d)
  con <- file(file.path(d, "bom.csv"), "wb")
  writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
  writeBin(charToRaw("name,label\n1,Caf\xC3\xA9\n"), con)
  close(con)
  survey <- fdf(type = "select_one_from_file bom.csv", name = "q", label = "Q")
  f <- read_quiet(make_form(survey), media_root = d)
  ch <- .file_choices(f, "bom.csv")$choices
  expect_equal(ch$name, "1")
  expect_equal(ch$label, "Caf\u00e9")
})

test_that("missing, empty and unsupported files", {
  skip_if_no_writexl()
  d <- write_csv_dir("list.xml" = "<root/>", "empty.csv" = "",
                     "places.geojson" = "{}")
  survey <- fdf(type = "text", name = "q", label = "Q")
  f <- read_quiet(make_form(survey), media_root = d)

  miss <- .file_choices(f, "nope.csv")
  expect_false(miss$found)
  expect_equal(nrow(miss$choices), 0L)
  expect_true(is.na(miss$path))

  # an XML file without items is like an empty CSV file
  xml <- .file_choices(f, "list.xml")
  expect_false(xml$found)
  expect_equal(nrow(xml$choices), 0L)
  expect_match(xml$note, "could not be read")

  geo <- .file_choices(f, "places.geojson")
  expect_true(geo$found)
  expect_equal(nrow(geo$choices), 0L)
  expect_match(geo$note, "CSV and XML")

  empty <- .file_choices(f, "empty.csv")
  expect_false(empty$found)
  expect_equal(nrow(empty$choices), 0L)

  expect_false(.file_choices(f, NA_character_)$found)
})

test_that("a file is found by basename and is not confused with a folder", {
  skip_if_no_writexl()
  d <- write_csv_dir("a.csv" = c("name,label", "1,One"))
  dir.create(file.path(d, "folder.csv"))
  f <- read_quiet(make_form(fdf(type = "text", name = "q", label = "Q")),
                  media_root = d)
  expect_true(.file_choices(f, "somewhere/else/a.csv")$found)
  expect_false(.file_choices(f, "folder.csv")$found)
})

test_that("select questions report their source", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  f <- read_quiet(p)
  idx <- function(nm) which(f$survey$name == nm)
  sc <- .select_choices(f, idx("water_source"))
  expect_equal(sc$source, "choices")
  expect_true(sc$source_found)
  expect_equal(sc$n_total, 4L)

  vf <- .select_choices(f, idx("village_file"))
  expect_equal(vf$source, "villages.csv")
  expect_true(vf$source_found)
  expect_equal(vf$n_total, 5L)

  mf <- .select_choices(f, idx("missing_items"))
  expect_equal(mf$source, "missing_list.csv")
  expect_true(mf$source_found)
  expect_equal(nrow(mf$choices), 3L)
})

test_that("a choice file that does not exist is reported as not found", {
  skip_if_no_writexl()
  f <- read_quiet(make_form(
    fdf(type = "select_one_from_file nope.csv", name = "q", label = "Q")
  ))
  mf <- .select_choices(f, 1L)
  expect_equal(mf$source, "nope.csv")
  expect_false(mf$source_found)
  expect_equal(nrow(mf$choices), 0L)
})

test_that("select_choices keeps every choice", {
  skip_if_no_writexl()
  f <- small_form()
  i <- which(f$survey$name == "assets")
  sc <- .select_choices(f, i)
  expect_equal(sc$choices$name, c("radio", "tv", "bike"))
  expect_equal(sc$n_total, 3L)
  expect_null(sc$truncated)
})

test_that("layout thresholds", {
  lab <- function(n, w = 5) rep(strrep("x", w), n)
  l <- function(n, w = 5, app = NA, mx = 30) {
    .choice_layout(lab(n, w), app, mx)
  }
  expect_equal(l(0)$layout, "list")
  expect_equal(l(1)$layout, "list")
  expect_equal(l(7)$layout, "list")
  expect_equal(l(7)$ncol, 1L)
  expect_equal(l(8)$layout, "columns")
  expect_equal(l(30)$layout, "columns")
  expect_equal(l(31)$layout, "code_box")
  # columns: 3 when every label has at most 25 characters, else 2
  expect_equal(l(8, 25)$ncol, 3L)
  expect_equal(l(8, 26)$ncol, 2L)
  expect_equal(l(31, 26)$ncol, 2L)
  # a custom limit
  expect_equal(l(12, mx = 10)$layout, "code_box")
  expect_equal(l(10, mx = 10)$layout, "columns")
  expect_equal(l(5000, mx = Inf)$layout, "columns")
  # minimal and autocomplete only matter above 7 choices
  expect_equal(l(8, app = "minimal")$layout, "code_box")
  expect_equal(l(8, app = "search('x') autocomplete")$layout, "code_box")
  expect_equal(l(7, app = "minimal")$layout, "list")
  expect_equal(l(8, app = "horizontal")$layout, "columns")
})

test_that("range values come from the parameters", {
  r <- .range_spec(list(start = 0, end = 10, step = 1))
  expect_equal(r$values, 0:10)
  expect_equal(c(r$start, r$end, r$step), c(0, 10, 1))
  # defaults: start 1, end 10, step 1
  d <- .range_spec(list())
  expect_equal(c(d$start, d$end, d$step), c(1, 10, 1))
  expect_equal(d$values, 1:10)
  # exactly 20 values is allowed, 21 is not
  expect_equal(length(.range_spec(list(start = 1, end = 20))$values), 20L)
  expect_null(.range_spec(list(start = 0, end = 100, step = 5))$values)
  expect_equal(.range_spec(list(start = 0, end = 100, step = 20))$values,
               seq(0, 100, 20))
  # decimals do not drift
  expect_equal(.range_spec(list(start = 0, end = 1, step = 0.1))$values,
               seq(0, 1, by = 0.1))
  # nonsense falls back to something usable
  expect_equal(.range_spec(list(step = 0))$step, 1)
  expect_equal(.range_spec(list(step = -2))$step, 1)
  expect_null(.range_spec(list(start = 5, end = 1))$values)
  expect_equal(.range_spec(list(start = "x"))$start, 1)
})
