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
  expect_equal(members$repeat_text, "Q3")
  expect_true(is.na(members$repeat_count_value))
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
  expect_equal(q$relevant_text, "Q1 is \"Yes\"")
  # the constraint message replaces the expression
  expect_equal(q$constraint_text, "Enter 0 to 30")
  expect_true(is.na(q$guidance_hint))
  expect_true(is.na(q$appearance))
  expect_true(is.na(q$choice_filter_text))
  expect_null(q$image)
  expect_false(q$answer$decimal)
  expect_equal(q$answer$range_text, "from 0 to 30")

  c1 <- block_named(m, "consent")
  expect_true(c1$required)
  expect_equal(c1$answer$symbol, "circle")
  expect_equal(c1$answer$choices$name, c("yes", "no"))
  expect_equal(c1$answer$n_total, 2L)
  expect_equal(c1$answer$layout, "list")
  expect_equal(c1$answer$choices$write_in, c(FALSE, FALSE))
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
               "Q3 > 1 or Q1 is \"No\"")
  # a child with its own relevance is combined, parenthesizing an "or"
  survey <- fdf(
    type = c("integer", "begin_group", "text", "end_group"),
    name = c("n", "g", "q", NA),
    label = c("N", NA, "Q", NA),
    relevant = c(NA, "${n} > 1 or ${n} < 0", "${n} != 5", NA)
  )
  m2 <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_equal(block_named(m2, "q")$relevant_text,
               "(Q1 > 1 or Q1 < 0) and Q1 is not 5")
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
  expect_equal(block_named(no_rel, "members", "section")$repeat_text, "Q3")
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
               paste0("Q2 does not include \"Radio\" and ",
                      "(Q3 > 2 or Q1 is \"Yes\")"))
  expect_equal(block_named(m, "members", "section")$repeat_text, "Q3")
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
               c("Latitude", "Longitude", "Accuracy (m)"))
  expect_equal(a("location")$subtype, "geopoint")
  np <- .cxf_strings("en")$not_on_paper
  expect_equal(a("trace_route")$prompt, np)
  expect_equal(a("trace_route")$reference, "Reference no.")
  expect_equal(a("plot_shape")$subtype, "geoshape")
  expect_equal(a("plot_shape")$prompt, np)
  expect_equal(a("photo_dwelling")$subtype, "image")
  for (nm in c("photo_dwelling", "audio_comment", "video_water",
               "consent_form")) {
    expect_equal(a(nm)$prompt, np, info = nm)
    expect_equal(a(nm)$reference, "Reference no.", info = nm)
  }
  # a range with more than 20 values is described in words
  expect_equal(a("food_share")$range_text, "0 to 100, steps of 5")
  expect_true(is.na(a("water_hours")$range_text))
  expect_equal(a("asset_tag")$lines, 1L)
  expect_equal(a("ack_understood")$prompt, "OK")

  # choices from a file (a missing file is covered in test-media.R)
  vf <- a("village_file")
  expect_equal(vf$source, "villages.csv")
  expect_true(vf$source_found)
  expect_equal(nrow(vf$choices), 5L)
  mi <- a("missing_items")
  expect_true(mi$source_found)
  expect_equal(nrow(mi$choices), 3L)

  # choice filter text and media records
  expect_equal(block_named(m, "village")$choice_filter_text,
               "district = Q12")
  img <- block_named(m, "intro_note")$image
  expect_true(img$found)
  expect_equal(img$file, "household.png")
  expect_true(img$width_in <= test_opts()$text_width_in && img$height_in <= 2)
  loc <- block_named(m, "location")$image
  expect_true(loc$found)
  # choice images are records inside a list column
  ws <- a("water_source")$choices
  expect_type(ws$image, "list")
  # four choices and the "other" row
  expect_equal(nrow(ws), 5L)
  expect_equal(ws$write_in, c(FALSE, FALSE, FALSE, FALSE, TRUE))
})

test_that("max_choices no longer truncates; it sets the layout", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  f <- read_quiet(p)
  m <- .build_model(f, test_opts())
  crop <- block_named(m, "main_crop")$answer
  expect_equal(nrow(crop$choices), 30L)
  expect_equal(crop$n_total, 30L)
  expect_equal(crop$layout, "columns")
  expect_null(crop$truncated)
  m10 <- .build_model(f, test_opts(max_choices = 10))
  crop10 <- block_named(m10, "main_crop")$answer
  expect_equal(nrow(crop10$choices), 30L)
  expect_equal(crop10$layout, "code_box")
  expect_true(crop10$ncol %in% 2:3)
  expect_equal(block_named(m10, "consent")$answer$layout, "list")
})

test_that("layouts follow the number of choices and the appearance", {
  skip_if_no_writexl()
  mk <- function(n, appearance = NA, label = "Item", or_other = FALSE) {
    survey <- fdf(type = paste0("select_one l", if (or_other) " or_other" else ""),
                  name = "q", label = "Q", appearance = appearance)
    choices <- fdf(list_name = rep("l", n), name = paste0("c", seq_len(n)),
                   label = paste(label, seq_len(n)))
    ans <- block_named(.build_model(read_quiet(make_form(survey, choices)),
                                    test_opts()), "q")$answer
    ans
  }
  expect_equal(mk(7)$layout, "list")
  expect_equal(mk(8)$layout, "columns")
  expect_equal(mk(8)$ncol, 3L)
  # one label above 25 characters: two columns
  expect_equal(mk(8, label = "A rather long choice label")$ncol, 2L)
  expect_equal(mk(30)$layout, "columns")
  expect_equal(mk(31)$layout, "code_box")
  expect_equal(mk(8, "minimal")$layout, "code_box")
  expect_equal(mk(7, "minimal")$layout, "list")
  expect_equal(mk(12, "autocomplete")$layout, "code_box")
})

test_that("or_other appends a write-in row", {
  skip_if_no_writexl()
  survey <- fdf(type = "select_one l or_other", name = "q", label = "Q")
  choices <- fdf(list_name = c("l", "l"), name = c("a", "b"),
                 label = c("A", "B"))
  f <- read_quiet(make_form(survey, choices))
  a <- block_named(.build_model(f, test_opts()), "q")$answer
  expect_true(a$or_other)
  expect_equal(a$choices$name, c("a", "b", "other"))
  expect_equal(a$choices$label, c("A", "B", "Other (specify)"))
  expect_equal(a$choices$write_in, c(FALSE, FALSE, TRUE))
  expect_equal(a$n_total, 2L)
  expect_type(a$choices$image, "list")
  expect_null(a$choices$image[[3]])
  # the label comes from the string table, and can be overridden
  fr <- block_named(.build_model(f, test_opts(strings = list(
    other_specify = "Autre"))), "q")$answer
  expect_equal(fr$choices$label[3], "Autre")
  # the extra row counts towards the layout (7 choices + other = 8)
  survey7 <- fdf(type = "select_one l7 or_other", name = "q", label = "Q")
  ch7 <- fdf(list_name = rep("l7", 7), name = letters[1:7],
             label = LETTERS[1:7])
  a7 <- block_named(.build_model(read_quiet(make_form(survey7, ch7)),
                                 test_opts()), "q")$answer
  expect_equal(a7$layout, "columns")
})

test_that("references in labels and hints become [Qn] or [name]", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c("integer", "begin_group", "text", "end_group", "calculate",
             "note"),
    name = c("age", "g", "q", NA, "calc", "n"),
    label = c("Age", "Group", "Is ${age} years right? See ${calc} and ${zzz}",
              NA, NA, "Note on ${ q }"),
    hint = c(NA, NA, "Compare with ${age}", NA, NA, NA),
    guidance_hint = c(NA, NA, "Check ${g}", NA, NA, NA)
  )
  m <- .build_model(read_quiet(make_form(survey)), test_opts())
  q <- block_named(m, "q")
  expect_equal(q$label, "Is [Q1] years right? See [calc] and [zzz]")
  expect_equal(q$hint, "Compare with [Q1]")
  expect_equal(q$guidance_hint, "Check [Q2]")
  expect_equal(block_named(m, "n")$label, "Note on [Q2.1]")
  nh <- .build_model(read_quiet(make_form(survey)),
                     test_opts(include_hints = FALSE))
  expect_true(is.na(block_named(nh, "q")$guidance_hint))
})

test_that("line breaks are normalized to \\n", {
  skip_if_no_writexl()
  survey <- fdf(type = c("text", "note"), name = c("q", "n"),
                label = c("Line one\r\nLine two\rLine three", "A\nB"),
                hint = c("H1\r\nH2", NA))
  m <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_equal(block_named(m, "q")$label, "Line one\nLine two\nLine three")
  expect_equal(block_named(m, "q")$hint, "H1\nH2")
  expect_equal(block_named(m, "n")$label, "A\nB")
})

test_that("appearance is carried on question blocks", {
  skip_if_no_writexl()
  survey <- fdf(type = "text", name = "q", label = "Q",
                appearance = "multiline")
  m <- .build_model(read_quiet(make_form(survey)), test_opts())
  expect_equal(block_named(m, "q")$appearance, "multiline")
})

test_that("show_variable_names controls references in logic text", {
  skip_if_no_writexl()
  f <- small_form()
  on <- .build_model(f, test_opts(show_variable_names = TRUE))
  off <- .build_model(f, test_opts(show_variable_names = FALSE))
  expect_equal(block_named(on, "hh_size")$relevant_text,
               "Q1 (consent) is \"Yes\"")
  expect_equal(block_named(off, "hh_size")$relevant_text, "Q1 is \"Yes\"")
  expect_equal(block_named(on, "members", "section")$repeat_text,
               "Q3 (hh_size)")
  expect_equal(block_named(on, "t1")$relevant_text,
               paste0("Q2 (assets) does not include \"Radio\" and ",
                      "(Q3 (hh_size) > 2 or Q1 (consent) is \"Yes\")"))
})

test_that("constraint text: message, humanized, or dropped", {
  skip_if_no_writexl()
  survey <- fdf(
    type = rep("integer", 4), name = c("a", "b", "c", "d"),
    label = c("A", "B", "C", "D"),
    constraint = c(". >= 0", ". >= 0", "regex(., '^[0-9]+$')",
                   "regex(., '^[0-9]+$')"),
    constraint_message = c(NA, "Must be positive", NA, "Digits only")
  )
  f <- read_quiet(make_form(survey))
  off <- .build_model(f, test_opts(show_variable_names = FALSE))
  on <- .build_model(f, test_opts(show_variable_names = TRUE))
  ct <- function(m, nm) block_named(m, nm)$constraint_text
  expect_equal(ct(off, "a"), "the answer >= 0")
  expect_equal(ct(off, "b"), "Must be positive")
  # verbatim fallback without a message: dropped unless names are shown
  expect_true(is.na(ct(off, "c")))
  expect_equal(ct(on, "c"), "regex(., '^[0-9]+$')")
  expect_equal(ct(off, "d"), "Digits only")
  # the range text still comes from the expression
  expect_equal(block_named(off, "b")$answer$range_text, "at least 0")
})

test_that("verbatim relevant expressions are kept", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c("integer", "text"), name = c("n", "q"), label = c("N", "Q"),
    relevant = c(NA, "string-length(${n}) > 3")
  )
  f <- read_quiet(make_form(survey))
  off <- .build_model(f, test_opts(show_variable_names = FALSE))
  expect_equal(block_named(off, "q")$relevant_text, "string-length(Q1) > 3")
})

test_that("repeat_count_value is numeric for constant counts", {
  skip_if_no_writexl()
  survey <- fdf(
    type = c("begin_repeat", "text", "end_repeat", "begin_repeat", "text",
             "end_repeat", "begin_repeat", "text", "end_repeat"),
    name = c("r1", "q1", NA, "r2", "q2", NA, "r3", "q3", NA),
    label = c("R1", "Q1", NA, "R2", "Q2", NA, "R3", "Q3", NA),
    repeat_count = c("3", NA, NA, "${q1}", NA, NA, NA, NA, NA)
  )
  m <- .build_model(read_quiet(make_form(survey)), test_opts())
  v <- function(nm) block_named(m, nm, "section")$repeat_count_value
  expect_equal(v("r1"), 3)
  expect_true(is.na(v("r2")))
  expect_true(is.na(v("r3")))
})

test_that("the legend explains rank questions and follows the language", {
  skip_if_no_writexl()
  sym <- .cxf_symbols()
  survey <- fdf(type = "rank l", name = "q", label = "Q")
  choices <- fdf(list_name = c("l", "l"), name = c("a", "b"),
                 label = c("A", "B"))
  m <- .build_model(read_quiet(make_form(survey, choices)), test_opts())
  expect_true(any(startsWith(m$meta$legend, sym$rank)))
  expect_true(any(grepl("Rank the answers", m$meta$legend)))
  ml <- .build_model(read_quiet(make_form(survey, choices)),
                     test_opts(strings = list(legend_rank = "Classez")))
  expect_true(any(grepl("Classez", ml$meta$legend)))
})

test_that("meta carries the strings, the font and rtl", {
  skip_if_no_writexl()
  m <- .build_model(small_form(), test_opts())
  expect_equal(m$meta$strings, .cxf_strings("en"))
  expect_equal(m$meta$font_family_cs, "Arial")
  expect_false(m$meta$rtl)
  m2 <- .build_model(small_form(), test_opts(font_family_cs = "My CS Font",
                                             font_family = "Calibri"))
  expect_equal(m2$meta$font_family_cs, "My CS Font")
  m3 <- .build_model(small_form(), test_opts(font_family = "Calibri"))
  expect_equal(m3$meta$font_family_cs, "Calibri")
})

test_that("script detection chooses the complex-script font", {
  d <- function(x) .detect_script_font(x, "Arial")
  expect_equal(d("Hello")$font_family_cs, "Arial")
  expect_equal(d(c(NA, ""))$font_family_cs, "Arial")
  expect_equal(d("\u0995\u09a5\u09be")$font_family_cs, "Noto Sans Bengali")
  expect_false(d("\u0995\u09a5\u09be")$rtl)
  ar <- d("\u0645\u0631\u062d\u0628\u0627")
  expect_equal(ar$font_family_cs, "Noto Naskh Arabic")
  expect_true(ar$rtl)
  expect_equal(d("\u0928\u092e\u0938\u094d\u0924\u0947")$font_family_cs,
               "Noto Sans Devanagari")
  expect_equal(d("\u1230\u120b\u121d")$font_family_cs, "Noto Sans Ethiopic")
  he <- d("\u05e9\u05dc\u05d5\u05dd")
  expect_equal(he$font_family_cs, "Arial")
  expect_true(he$rtl)
  # the script with most characters wins
  expect_equal(d(c("\u0995\u09a5\u09be\u0995", "\u0645"))$font_family_cs,
               "Noto Sans Bengali")
})

test_that("the Bengali example gets the Bengali font", {
  p <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  bn <- .build_model(read_quiet(p, lang = "bn"), test_opts())
  expect_equal(bn$meta$font_family_cs, "Noto Sans Bengali")
  expect_false(bn$meta$rtl)
  en <- .build_model(read_quiet(p, lang = "en"), test_opts())
  expect_equal(en$meta$font_family_cs, "Arial")
})

test_that("French forms use French strings", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  fr <- .build_model(read_quiet(p, lang = "fr"), test_opts())
  expect_equal(fr$meta$strings, .cxf_strings("fr"))
  expect_true(any(grepl("Question obligatoire", fr$meta$legend)))
  other <- block_named(fr, "water_source")$answer$choices
  expect_equal(other$label[nrow(other)], "Autre (pr\u00e9ciser)")
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
