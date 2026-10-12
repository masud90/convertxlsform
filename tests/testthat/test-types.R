test_that("parse_type normalizes every type", {
  x <- c("begin group", "begin_group", "End Group", "begin repeat",
         "end_repeat", "dateTime", "select one yesno", "select_one yesno",
         "select multiple fruit", "select_multiple fruit",
         "select_one_from_file x.csv", "select_multiple_from_file y.csv",
         "select_one_external cities", "rank feature", "  text  ",
         "start-geopoint", "background-audio", "integer", NA, "")
  p <- .parse_type(x)
  expect_equal(
    p$base_type,
    c("begin_group", "begin_group", "end_group", "begin_repeat", "end_repeat",
      "datetime", "select_one", "select_one", "select_multiple",
      "select_multiple", "select_one_from_file", "select_multiple_from_file",
      "select_one_external", "rank", "text", "start-geopoint",
      "background-audio", "integer", NA, NA)
  )
  expect_equal(p$type_arg[c(7, 8, 9, 10, 11, 12, 13, 14, 15)],
               c("yesno", "yesno", "fruit", "fruit", "x.csv", "y.csv",
                 "cities", "feature", NA))
  expect_false(any(p$or_other))
})

test_that("parse_type handles or_other and keeps list-name case", {
  p <- .parse_type(c("select_one Colours or_other",
                     "select multiple Fruit or other", "select_one a"))
  expect_equal(p$or_other, c(TRUE, TRUE, FALSE))
  expect_equal(p$type_arg, c("Colours", "Fruit", "a"))
  expect_equal(p$base_type, c("select_one", "select_multiple", "select_one"))
})

test_that("classify_types covers the registry and flags unknowns", {
  reg <- .type_registry()
  expect_equal(.classify_types(reg$base_type), reg$kind)
  expect_equal(
    .classify_types(c("integer", "decimal", "geotrace", "trigger", "calculate",
                      "osm", NA)),
    c("number", "number", "geo", "acknowledge", "hidden", "unknown", "unknown")
  )
  expect_equal(.type_subtype(c("integer", "decimal", "geoshape", "text")),
               c("integer", "decimal", "geoshape", NA))
  for (k in c("select_one", "select_multiple", "rank", "text", "number",
              "range", "date", "time", "datetime", "geo", "media", "barcode",
              "acknowledge", "note", "hidden", "begin_group", "end_group",
              "begin_repeat", "end_repeat")) {
    expect_true(k %in% reg$kind)
  }
})

test_that("parse_parameters handles separators and numbers", {
  p <- .parse_parameters(c("start=1 end=10 step=1", "start=0, end=5; step=0.5",
                           "randomize=true seed=42", "start = 2   end = 8",
                           NA, "", "   ", "junk", "Start=-1"))
  expect_equal(p[[1]], list(start = 1, end = 10, step = 1))
  expect_equal(p[[2]], list(start = 0, end = 5, step = 0.5))
  expect_equal(p[[3]], list(randomize = "true", seed = 42))
  expect_equal(p[[4]], list(start = 2, end = 8))
  for (i in 5:8) expect_length(p[[i]], 0)
  expect_equal(p[[9]], list(start = -1))
  expect_named(p[[5]], character(0))
})

test_that("parameters keep the letter s inside values", {
  p <- .parse_parameters("randomize=true seed=1 label=sums")
  expect_equal(p[[1]]$label, "sums")
})

test_that("parse_required distinguishes flags and expressions", {
  r <- .parse_required(c("yes", "TRUE", "true()", "1", "Yes", NA, "", "no",
                         "false()", "0", "${a} = 'x'"))
  expect_equal(r$required,
               c(TRUE, TRUE, TRUE, TRUE, TRUE, FALSE, FALSE, FALSE, FALSE,
                 FALSE, TRUE))
  expect_equal(r$required_expr, c(rep(NA, 10), "${a} = 'x'"))
})
