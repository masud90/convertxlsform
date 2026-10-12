build <- function(type, ...) {
  convertxlsform:::.build_structure(make_struct_survey(type, ...), verbose = FALSE)
}

test_that("flat form is numbered 1, 2, 3", {
  out <- build(c("text", "integer", "note"))
  expect_equal(out$serial, c("1", "2", "3"))
  expect_true(all(out$is_question))
  expect_equal(out$depth, c(0L, 0L, 0L))
})

test_that("nested numbering inside a group gives 1.1, 1.2 and continues at top level", {
  out <- build(c("begin_group", "text", "integer", "end_group", "text"))
  expect_equal(out$kind, c("begin_group", "question", "question", "end_group", "question"))
  expect_equal(out$serial[out$is_question], c("1.1", "1.2", "2"))
})

test_that("group row itself carries no serial and depth is tracked", {
  out <- build(c("begin_group", "text", "end_group"))
  expect_true(is.na(out$serial[1]))
  expect_equal(out$depth, c(1L, 1L, 0L))
})

test_that("deeper nesting gives three-level serials", {
  out <- build(c("begin_group", "begin_repeat", "text", "text", "end_repeat", "text", "end_group"))
  expect_equal(out$serial[out$is_question], c("1.1.1", "1.1.2", "1.2"))
})

test_that("calculate/hidden rows are ignored and unnumbered", {
  out <- build(c("text", "calculate", "text"))
  expect_equal(out$kind, c("question", "ignored", "question"))
  expect_equal(out$serial, c("1", NA, "2"))
})

test_that("numbering off gives NA serials", {
  s <- make_struct_survey(c("begin_group", "text", "text", "end_group"))
  out <- convertxlsform:::.build_structure(s, nested_serial_numbers = FALSE, verbose = FALSE)
  expect_true(all(is.na(out$serial)))
  expect_equal(sum(out$is_question), 2L)
})

test_that("unclosed begin_group errors", {
  s <- make_struct_survey(c("begin_group", "text"))
  expect_error(convertxlsform:::.build_structure(s, verbose = FALSE),
               "Unclosed begin_group starting at survey row 1")
})

test_that("end without begin errors", {
  s <- make_struct_survey(c("text", "end_group"))
  expect_error(convertxlsform:::.build_structure(s, verbose = FALSE),
               "Unmatched group terminator at survey row 2")
})

test_that("mismatched end marker errors", {
  s <- make_struct_survey(c("begin_group", "text", "end_repeat"))
  expect_error(convertxlsform:::.build_structure(s, verbose = FALSE),
               "Mismatched end marker at survey row 3")
})

test_that("structure of the basic fixture via the full reader", {
  path <- make_basic_xlsform()
  ph2 <- convertxlsform:::.read_and_validate_xlsform(path, verbose = FALSE)
  out <- convertxlsform:::.build_structure(ph2$survey, verbose = FALSE)
  expect_equal(nrow(out), 4L)
  expect_equal(out$serial[out$is_question], c("1.1", "1.2"))
})

test_that("verbose structure build reports counts", {
  s <- make_struct_survey(c("begin_repeat", "text", "end_repeat"))
  expect_message(convertxlsform:::.build_structure(s, verbose = TRUE), "begin_repeat=1")
})
