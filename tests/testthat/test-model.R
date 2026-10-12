test_that("blocks come out in form order with the expected types", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  expect_s3_class(m, "cxf_model")
  expect_named(m, c("meta", "blocks"))
  expect_equal(
    block_types(m),
    c("question", "question", "question", "section", "question", "question",
      "section", "question", "section_end", "note")
  )
  nm <- vapply(m$blocks, function(b) b$name %||% NA_character_, "")
  expect_equal(nm, c("consent", "assets", "hh_size", "g1", "t1", "t2",
                     "members", "mname", "members", "n1"))
  ser <- vapply(m$blocks, function(b) b$serial, "")
  expect_equal(ser, c("1", "2", "3", "4", "4.1", "4.2", "5", "5.1", "5",
                      NA))
  # every block has the common fields
  for (b in m$blocks) {
    expect_true(all(c("block", "row", "depth", "serial", "name") %in% names(b)))
  }
  # the calculate row is skipped by default
  expect_null(block_named(m, "calc"))
})

test_that("meta carries title, form id, version, language, legend", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  expect_equal(m$meta$title, "Small form")
  expect_equal(m$meta$form_id, "small")
  expect_equal(m$meta$version, "1")
  expect_true(is.na(m$meta$language))
  expect_s3_class(m$meta$generated, "Date")
  expect_type(m$meta$legend, "character")
})

test_that("level counts visible sections, depth counts every group", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  g1 <- block_named(m, "g1")
  t1 <- block_named(m, "t1")
  t2 <- block_named(m, "t2")
  members <- block_named(m, "members", "section")
  mname <- block_named(m, "mname")
  # t2 sits in two groups, one of them transparent
  expect_equal(c(t1$depth, t2$depth), c(1L, 2L))
  expect_equal(c(t1$level, t2$level), c(1L, 1L))
  expect_equal(g1$level, 1L)
  expect_equal(members$level, 1L)
  expect_equal(mname$level, 1L)
  expect_equal(block_named(m, "consent")$level, 0L)
  expect_equal(block_named(m, "members", "section_end")$level, 1L)
  expect_equal(block_named(m, "n1")$level, 0L)
  # the transparent group has no block of its own
  expect_null(block_named(m, "g_hidden"))
})

test_that("section levels nest and are capped at 3", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c(rep("begin_group", 5), "text", rep("end_group", 5), "text"),
    name = c("a", "b", "c", "d", "e", "q", rep(NA, 5), "last"),
    label = c("A", "B", "C", "D", "E", "Q", rep(NA, 5), "Last")
  )
  f <- read_quiet(make_form(survey))
  m <- .build_model(f, test_opts())
  lv <- function(nm) block_named(m, nm)$level
  expect_equal(c(lv("a"), lv("b"), lv("c"), lv("d"), lv("e")),
               c(1L, 2L, 3L, 3L, 3L))
  expect_equal(lv("q"), 3L)
  expect_equal(lv("last"), 0L)
  expect_equal(block_named(m, "q")$depth, 5L)
})

test_that("a labelless repeat is still a section", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c("begin_repeat", "text", "end_repeat"),
    name = c("r", "q", NA),
    label = c(NA, "Q", NA),
    repeat_count = c("3", NA, NA)
  )
  m <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_equal(block_types(m), c("section", "question", "section_end"))
  s <- m$blocks[[1]]
  expect_true(s$is_repeat)
  expect_equal(s$repeat_text, "3")
  expect_equal(s$serial, "1")
  expect_equal(m$blocks[[2]]$serial, "1.1")
})

test_that("section and question fields", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  members <- block_named(m, "members", "section")
  expect_true(members$is_repeat)
  expect_equal(members$repeat_text, "Q3 (hh_size)")
  expect_equal(members$label, "Members")
  g1 <- block_named(m, "g1")
  expect_false(g1$is_repeat)
  expect_true(is.na(g1$repeat_text))

  q <- block_named(m, "hh_size")
  expect_equal(q$block, "question")
  expect_equal(q$kind, "number")
  expect_equal(q$subtype, "integer")
  expect_false(q$required)
  expect_equal(q$hint, "Count everyone")
  expect_equal(q$relevant_text, "Q1 (consent) is \"Yes\"")
  expect_equal(q$constraint_text,
               "the answer >= 0 and the answer <= 30 (Enter 0 to 30)")
  expect_true(is.na(q$choice_filter_text))
  expect_null(q$image)
  expect_false(q$answer$decimal)
  expect_equal(q$answer$range_text, "from 0 to 30")

  c1 <- block_named(m, "consent")
  expect_true(c1$required)
  expect_equal(c1$answer$symbol, "circle")
  expect_equal(c1$answer$choices$name, c("yes", "no"))
  expect_equal(c1$answer$n_total, 2L)
  expect_false(c1$answer$truncated)
  expect_equal(c1$answer$source, "choices")
  expect_true(c1$answer$source_found)
  expect_false(c1$answer$or_other)
  expect_equal(block_named(m, "assets")$answer$symbol, "square")

  n1 <- block_named(m, "n1")
  expect_equal(n1$block, "note")
  expect_equal(n1$label, "Thank you")
})

test_that("a hidden group's relevance passes to its children", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  t2 <- block_named(m, "t2")
  expect_equal(t2$relevant_text,
               "Q3 (hh_size) > 1 or Q1 (consent) is \"No\"")
  # a child with its own relevance is combined, parenthesizing an "or"
  survey <- fdf(
    type = c("integer", "begin_group", "text", "end_group"),
    name = c("n", "g", "q", NA),
    label = c("N", NA, "Q", NA),
    relevant = c(NA, "${n} > 1 or ${n} < 0", "${n} != 5", NA)
  )
  m2 <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_equal(block_named(m2, "q")$relevant_text,
               "(Q1 (n) > 1 or Q1 (n) < 0) and Q1 (n) is not 5")
  # off with include_relevants = FALSE
  m3 <- .build_model(read_quiet(make_form(survey)),
                     test_opts(include_relevants = FALSE))
  expect_true(is.na(block_named(m3, "q")$relevant_text))
})

test_that("include_relevants, include_constraints and include_hints toggle text", {
  skip_if_no_writexl()
  f <- small_form()
  on <- .build_model(f, test_opts())
  expect_false(is.na(block_named(on, "t1")$relevant_text))
  expect_false(is.na(block_named(on, "hh_size")$constraint_text))
  expect_false(is.na(block_named(on, "hh_size")$hint))

  no_rel <- .build_model(f, test_opts(include_relevants = FALSE))
  expect_true(is.na(block_named(no_rel, "t1")$relevant_text))
  expect_true(is.na(block_named(no_rel, "hh_size")$relevant_text))
  expect_false(is.na(block_named(no_rel, "hh_size")$constraint_text))
  # repeat_count is not skip logic
  expect_equal(block_named(no_rel, "members", "section")$repeat_text,
               "Q3 (hh_size)")
  expect_false(grepl("Ask if", paste(no_rel$meta$legend, collapse = "|")))
  expect_true(grepl("Ask if", paste(on$meta$legend, collapse = "|")))

  no_con <- .build_model(f, test_opts(include_constraints = FALSE))
  expect_true(is.na(block_named(no_con, "hh_size")$constraint_text))
  expect_false(is.na(block_named(no_con, "hh_size")$relevant_text))

  no_hint <- .build_model(f, test_opts(include_hints = FALSE))
  expect_true(is.na(block_named(no_hint, "hh_size")$hint))
  expect_false(is.na(block_named(no_hint, "hh_size")$constraint_text))
})

test_that("include_hidden adds hidden fields as question blocks", {
  skip_if_no_writexl()
  f <- small_form()
  off <- .build_model(f, test_opts())
  on <- .build_model(f, test_opts(include_hidden = TRUE))
  expect_equal(length(on$blocks), length(off$blocks) + 1L)
  calc <- block_named(on, "calc")
  expect_equal(calc$block, "question")
  expect_equal(calc$kind, "hidden")
  expect_equal(calc$subtype, "calculate")
  expect_equal(calc$calculation, "1 + 1")
  expect_equal(calc$answer$calculation, "1 + 1")
  expect_true(is.na(calc$serial))
  # shown in the order of the form, between the repeat and the closing note
  expect_equal(tail(block_types(on), 2), c("question", "note"))
})

test_that("title override and fallbacks", {
  skip_if_no_writexl()
  f <- small_form()
  expect_equal(.build_model(f, test_opts(title = "My title"))$meta$title,
               "My title")
  # no title in settings: the file name
  survey <- fdf(type = "text", name = "q", label = "Q")
  path <- make_form(survey)
  m <- .build_model(read_quiet(path), test_opts())
  expect_equal(m$meta$title, tools::file_path_sans_ext(basename(path)))
})

test_that("flat numbering renumbers questions and the expression text", {
  skip_if_no_writexl()
  f <- small_form()
  m <- .build_model(f, test_opts(nested_serial_numbers = FALSE))
  ser <- vapply(m$blocks, function(b) b$serial, "")
  # groups are unnumbered, questions run 1..n
  expect_equal(ser, c("1", "2", "3", NA, "4", "5", NA, "6", NA, NA))
  expect_equal(block_named(m, "t1")$relevant_text,
               paste0("Q2 (assets) does not include \"Radio\" and ",
                      "(Q3 (hh_size) > 2 or Q1 (consent) is \"Yes\")"))
  expect_equal(block_named(m, "members", "section")$repeat_text, "Q3 (hh_size)")
  # groups and repeats are still sections
  expect_equal(sum(block_types(m) == "section"), 2L)
  # the input form is not modified
  expect_equal(f$survey$serial[which(f$survey$name == "t1")], "4.1")
})

test_that("the legend only explains what is on the form", {
  skip_if_no_writexl()
  sym <- .cxf_symbols()
  m <- .build_model(small_form(), test_opts())
  lg <- m$meta$legend
  expect_true(any(grepl("^\\* ", lg)))
  expect_true(any(startsWith(lg, sym$circle)))
  expect_true(any(startsWith(lg, sym$square)))
  expect_true(any(grepl("^Ask if", lg)))

  survey <- fdf(type = "text", name = "q", label = "Q")
  m2 <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_length(m2$meta$legend, 0L)
})

test_that("answer specs for every kind (all_types example)", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  m <- .build_model(read_quiet(p), test_opts())
  a <- function(nm) block_named(m, nm)$answer

  expect_equal(a("hh_code")$lines, 2L)
  expect_equal(a("comments")$lines, 4L)
  expect_true(a("land_ha")$decimal)
  expect_equal(a("visit_date")$pattern, "__ / __ / ____  (dd/mm/yyyy)")
  expect_equal(a("visit_time")$pattern, "__ : __  (hh:mm)")
  expect_equal(a("last_rain")$pattern,
               "__ / __ / ____  (dd/mm/yyyy)   __ : __  (hh:mm)")
  expect_equal(a("water_hours")$values, 0:10)
  expect_equal(c(a("food_share")$start, a("food_share")$end,
                 a("food_share")$step), c(0, 100, 5))
  expect_null(a("food_share")$values)
  expect_true(a("water_source")$or_other)
  expect_equal(a("priorities")$symbol, "rank")
  expect_equal(a("location")$fields,
               c("Latitude", "Longitude", "Altitude", "Accuracy"))
  expect_equal(a("location")$subtype, "geopoint")
  expect_equal(a("trace_route")$fields, "List of points")
  expect_equal(a("plot_shape")$subtype, "geoshape")
  expect_equal(a("photo_dwelling")$subtype, "image")
  expect_equal(a("photo_dwelling")$prompt, "[Photo to be taken]")
  expect_equal(a("audio_comment")$prompt, "[Audio to be recorded]")
  expect_equal(a("video_water")$prompt, "[Video to be recorded]")
  expect_equal(a("consent_form")$prompt, "[File to be attached]")
  expect_equal(a("asset_tag")$lines, 1L)
  expect_equal(a("ack_understood")$prompt, "OK")

  # choices from a file, and from a file that does not exist
  vf <- a("village_file")
  expect_equal(vf$source, "villages.csv")
  expect_true(vf$source_found)
  expect_equal(nrow(vf$choices), 5L)
  mi <- a("missing_items")
  expect_false(mi$source_found)
  expect_equal(nrow(mi$choices), 0L)

  # choice filter text and media records
  expect_equal(block_named(m, "village")$choice_filter_text,
               "district = Q12 (district)")
  img <- block_named(m, "intro_note")$image
  expect_true(img$found)
  expect_equal(img$file, "household.png")
  expect_true(img$width_in <= test_opts()$text_width_in && img$height_in <= 2)
  loc <- block_named(m, "location")$image
  expect_true(loc$found)
  # choice images are records inside a list column
  ws <- a("water_source")$choices
  expect_type(ws$image, "list")
  expect_equal(nrow(ws), 4L)
})

test_that("max_choices truncates long lists in the model", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  m <- .build_model(read_quiet(p), test_opts(max_choices = 10))
  crop <- block_named(m, "main_crop")$answer
  expect_equal(nrow(crop$choices), 10L)
  expect_equal(crop$n_total, 30L)
  expect_true(crop$truncated)
  expect_false(block_named(m, "consent")$answer$truncated)
})

test_that("language display names and choice labels follow lang", {
  p <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  en <- .build_model(read_quiet(p, lang = "en"), test_opts())
  bn <- .build_model(read_quiet(p, lang = "bn"), test_opts())
  expect_equal(en$meta$language, "English")
  expect_equal(bn$meta$language, "Bengali")
  expect_equal(length(en$blocks), length(bn$blocks))
  e <- block_named(en, "dwelling_ownership")$answer$choices$label
  b <- block_named(bn, "dwelling_ownership")$answer$choices$label
  expect_equal(e, c("Yes", "No"))
  expect_false(identical(e, b))
  # choice list from a file, language-aware
  ef <- block_named(en, "enumname")$answer$choices$label[1]
  bf <- block_named(bn, "enumname")$answer$choices$label[1]
  expect_equal(ef, "Grace Young")
  expect_false(identical(ef, bf))
})
