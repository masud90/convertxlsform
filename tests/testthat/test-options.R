test_that("defaults are complete and typed", {
  opts <- .cxf_options()
  expect_s3_class(opts, "cxf_options")
  expect_null(opts$lang)
  expect_true(opts$nested_serial_numbers)
  expect_true(opts$include_relevants)
  expect_true(opts$include_constraints)
  expect_true(opts$include_hints)
  expect_true(opts$include_choice_name)
  expect_false(opts$show_variable_names)
  expect_false(opts$include_hidden)
  expect_null(opts$media_root)
  expect_null(opts$title)
  expect_equal(opts$page_size, "A4")
  expect_equal(opts$orientation, "portrait")
  expect_equal(opts$margins_in, 1)
  expect_equal(opts$font_family, "Arial")
  expect_equal(opts$font_size, 11)
  expect_null(opts$template)
  expect_equal(opts$max_choices, 30)
  expect_null(opts$strings)
  expect_identical(opts$repeat_copies, 1L)
  expect_null(opts$font_family_cs)
  expect_false(opts$overwrite)
  expect_true(opts$verbose)
})

test_that("A4 portrait derives page and text width", {
  opts <- .cxf_options()
  expect_equal(opts$page_width_in, 8.27)
  expect_equal(opts$page_height_in, 11.69)
  expect_equal(opts$text_width_in, 6.27)
})

test_that("letter landscape swaps the page and derives text width", {
  opts <- .cxf_options(page_size = "letter", orientation = "landscape")
  expect_equal(opts$page_width_in, 11)
  expect_equal(opts$page_height_in, 8.5)
  expect_equal(opts$text_width_in, 9)
})

test_that("letter portrait and margins feed the text width", {
  opts <- .cxf_options(page_size = "letter", margins_in = 0.5)
  expect_equal(opts$page_width_in, 8.5)
  expect_equal(opts$page_height_in, 11)
  expect_equal(opts$text_width_in, 7.5)
})

test_that("page_size and orientation default to the first choice", {
  expect_equal(.cxf_options(page_size = c("A4", "letter"))$page_size, "A4")
  expect_equal(
    .cxf_options(orientation = c("portrait", "landscape"))$orientation,
    "portrait"
  )
})

test_that("unnamed and unknown options are rejected", {
  expect_error(.cxf_options(1), "named")
  expect_error(.cxf_options(not_an_option = 1), "Unknown option")
})

test_that("logical flags must be single TRUE or FALSE", {
  flags <- c("nested_serial_numbers", "include_relevants",
             "include_constraints", "include_hints", "include_choice_name",
             "show_variable_names", "include_hidden", "overwrite", "verbose")
  for (flag in flags) {
    args <- list(NA)
    names(args) <- flag
    expect_error(do.call(.cxf_options, args), flag, fixed = TRUE)
    args[[1]] <- "yes"
    expect_error(do.call(.cxf_options, args), flag, fixed = TRUE)
    args[[1]] <- c(TRUE, FALSE)
    expect_error(do.call(.cxf_options, args), flag, fixed = TRUE)
  }
})

test_that("margins_in must be a positive number of at most 2", {
  expect_error(.cxf_options(margins_in = 0), "margins_in")
  expect_error(.cxf_options(margins_in = 2.01), "margins_in")
  expect_error(.cxf_options(margins_in = -1), "margins_in")
  expect_error(.cxf_options(margins_in = "1"), "margins_in")
  expect_error(.cxf_options(margins_in = NA), "margins_in")
  expect_error(.cxf_options(margins_in = Inf), "margins_in")
  expect_equal(.cxf_options(margins_in = 2)$margins_in, 2)
})

test_that("font_size must be between 6 and 20", {
  expect_error(.cxf_options(font_size = 5), "font_size")
  expect_error(.cxf_options(font_size = 21), "font_size")
  expect_error(.cxf_options(font_size = "11"), "font_size")
  expect_equal(.cxf_options(font_size = 6)$font_size, 6)
  expect_equal(.cxf_options(font_size = 20)$font_size, 20)
})

test_that("max_choices must be a positive whole number or Inf", {
  expect_error(.cxf_options(max_choices = 2.5), "max_choices")
  expect_error(.cxf_options(max_choices = 0), "max_choices")
  expect_error(.cxf_options(max_choices = -3), "max_choices")
  expect_error(.cxf_options(max_choices = NA), "max_choices")
  expect_error(.cxf_options(max_choices = "10"), "max_choices")
  expect_equal(.cxf_options(max_choices = 10)$max_choices, 10)
  expect_equal(.cxf_options(max_choices = Inf)$max_choices, Inf)
})

test_that("page_size and orientation reject unknown values", {
  expect_error(.cxf_options(page_size = "B5"), "page_size")
  expect_error(.cxf_options(orientation = "square"), "orientation")
  expect_error(.cxf_options(page_size = 1), "page_size")
})

test_that("template must be an existing .docx file", {
  expect_error(.cxf_options(template = "reference.pdf"), ".docx")
  expect_error(
    .cxf_options(template = file.path(tempdir(), "missing-ref.docx")),
    "not found"
  )
  docx <- tempfile(fileext = ".docx")
  file.create(docx)
  on.exit(unlink(docx), add = TRUE)
  expect_equal(.cxf_options(template = docx)$template, docx)
})

test_that("string options reject blank and vector values", {
  expect_error(.cxf_options(lang = ""), "lang")
  expect_error(.cxf_options(lang = c("en", "fr")), "lang")
  expect_error(.cxf_options(title = 1), "title")
  expect_error(.cxf_options(media_root = NA_character_), "media_root")
  expect_error(.cxf_options(font_family = ""), "font_family")
  expect_equal(.cxf_options(lang = "en")$lang, "en")
})

test_that("strings must be a named list of known single strings", {
  ok <- .cxf_options(strings = list(other_specify = "Autre", ask_if = "Si"))
  expect_equal(ok$strings$other_specify, "Autre")
  expect_error(.cxf_options(strings = list(nope = "x")), "nope")
  expect_error(.cxf_options(strings = list(nope = "x")), "other_specify")
  expect_error(.cxf_options(strings = list("x")), "strings")
  expect_error(.cxf_options(strings = c(ask_if = "x")), "strings")
  expect_error(.cxf_options(strings = list(ask_if = c("a", "b"))), "strings")
  expect_error(.cxf_options(strings = list(ask_if = 1)), "strings")
  expect_error(.cxf_options(strings = list(ask_if = NA_character_)),
               "strings")
  expect_error(.cxf_options(strings = list()), "strings")
})

test_that("repeat_copies is a whole number from 1 to 20", {
  expect_identical(.cxf_options(repeat_copies = 3)$repeat_copies, 3L)
  expect_identical(.cxf_options(repeat_copies = 20L)$repeat_copies, 20L)
  expect_error(.cxf_options(repeat_copies = 0), "repeat_copies")
  expect_error(.cxf_options(repeat_copies = 21), "repeat_copies")
  expect_error(.cxf_options(repeat_copies = 1.5), "repeat_copies")
  expect_error(.cxf_options(repeat_copies = "2"), "repeat_copies")
  expect_error(.cxf_options(repeat_copies = NA), "repeat_copies")
})

test_that("font_family_cs is NULL or a string", {
  expect_equal(.cxf_options(font_family_cs = "Noto Sans")$font_family_cs,
               "Noto Sans")
  expect_error(.cxf_options(font_family_cs = ""), "font_family_cs")
  expect_error(.cxf_options(font_family_cs = 1), "font_family_cs")
})
