# Tests for forward skip instructions (R/skips.R).

skip_form <- function(...) {
  skip_if_no_writexl()
  sv <- fdf(
    type = c("select_one yn", "integer", "text", "begin_group", "text",
             "end_group", "text", "select_one yn", "text"),
    name = c("q1", "q2", "q3", "g", "q4", NA, "q5", "q6", "q7"),
    label = c("Q1", "Q2", "Q3", "G", "Q4", NA, "Q5", "Q6", "Q7"),
    relevant = c(NA, "${q1} = 'yes'", "selected(${q1}, 'yes')", NA, NA, NA,
                 NA, NA, "${q6} != 'no'")
  )
  ch <- fdf(list_name = c("yn", "yn"), name = c("yes", "no"),
            label = c("Yes", "No"))
  read_quiet(make_form(sv, ch), ...)
}

test_that(".sk_eval decides relevance from one answer", {
  ev <- function(expr, v) .sk_eval(.hx_parse(.hx_tokenize(expr)), "s", v)
  expect_true(ev("${s} = 'a'", "a"))
  expect_false(ev("${s} = 'a'", "b"))
  expect_true(ev("'a' = ${s}", "a"))
  expect_false(ev("${s} != 'a'", "a"))
  expect_true(ev("selected(${s}, 'b')", "b"))
  expect_false(ev("not(selected(${s}, 'b'))", "b"))
  expect_true(ev("${s} = 1", "1"))
  expect_true(ev("${s} > 2", "3"))
  expect_false(ev("${s} >= 2", "1"))
  expect_true(ev("2 < ${s}", "3"))
  # three-valued logic with other answers
  expect_false(ev("${s} = 'a' and ${t} = 'x'", "b"))
  expect_true(is.na(ev("${s} = 'a' and ${t} = 'x'", "a")))
  expect_true(ev("${s} = 'a' or ${t} = 'x'", "a"))
  expect_true(is.na(ev("${s} = 'a' or ${t} = 'x'", "b")))
  expect_true(is.na(ev("string-length(${s}) > 2", "abc")))
})

test_that("skip targets name the next asked question", {
  f <- skip_form()
  sk <- .skip_targets(f)
  expect_equal(unname(sk[[1L]]), c(NA, "4"))
  expect_named(sk[[1L]], c("yes", "no"))
  expect_equal(unname(sk[[8L]]), c(NA, "0"))
  expect_null(sk[[2L]])
})

test_that("no instruction when the next question is asked anyway", {
  f <- small_form()
  expect_null(.skip_targets(f)[[1L]])
})

test_that("unknown relevance stops the instruction", {
  skip_if_no_writexl()
  sv <- fdf(type = c("select_one yn", "text", "text"),
            name = c("q1", "q2", "q3"), label = c("Q1", "Q2", "Q3"),
            relevant = c(NA, "${q1} = 'yes' and ${q9} = 1", NA))
  ch <- fdf(list_name = c("yn", "yn"), name = c("yes", "no"),
            label = c("Yes", "No"))
  sk <- .skip_targets(read_quiet(make_form(sv, ch)))
  # "no" hides q2 whatever q9 is; "yes" depends on q9
  expect_equal(unname(sk[[1L]]), c(NA, "3"))
})

test_that("skips do not leave a repeat", {
  skip_if_no_writexl()
  sv <- fdf(type = c("begin_repeat", "select_one yn", "text", "end_repeat",
                     "text"),
            name = c("r", "q1", "q2", NA, "q3"),
            label = c("R", "Q1", "Q2", NA, "Q3"),
            relevant = c(NA, NA, "${q1} = 'yes'", NA, NA))
  ch <- fdf(list_name = c("yn", "yn"), name = c("yes", "no"),
            label = c("Yes", "No"))
  expect_null(.skip_targets(read_quiet(make_form(sv, ch)))[[2L]])
})

test_that("the model carries skip text and a legend line", {
  f <- skip_form()
  m <- .build_model(f, test_opts())
  q1 <- block_named(m, "q1")
  expect_equal(q1$answer$choices$skip, c(NA, "Go to Q4"))
  q6 <- block_named(m, "q6")
  expect_equal(q6$answer$choices$skip, c(NA, "Go to end"))
  expect_true(any(grepl("^→ Go to:", m$meta$legend)))
  off <- .build_model(f, test_opts(include_skips = FALSE))
  expect_true(all(is.na(block_named(off, "q1")$answer$choices$skip)))
  expect_false(any(grepl("Go to", off$meta$legend)))
})

test_that("flat numbering changes the target", {
  f <- skip_form()
  m <- .build_model(f, test_opts(nested_serial_numbers = FALSE))
  expect_equal(block_named(m, "q1")$answer$choices$skip, c(NA, "Go to Q4"))
})
