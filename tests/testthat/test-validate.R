test_that("missing survey sheet aborts and lists sheets", {
  skip_if_no_writexl()
  p <- make_form(fdf(type = "text", name = "a", label = "A"),
                 survey_name = "questions")
  expect_error(read_quiet(p), "Sheet.*survey.*not found")
  expect_error(read_quiet(p), "questions")
})

test_that("missing required columns are reported with sheet and column", {
  skip_if_no_writexl()
  p <- make_form(fdf(name = "a", label = "A"))
  expect_error(read_quiet(p), "Sheet `survey`, row 1: column `type` is missing")
  p <- make_form(fdf(type = "text", name = "a", other = "A"))
  expect_error(read_quiet(p), "no `label` column")
  p <- make_form(fdf(type = "select_one l", name = "a", label = "A"),
                 fdf(list_name = "l", label = "Yes"))
  expect_error(read_quiet(p), "Sheet `choices`, row 1: column `name`")
})

test_that("choices sheet is needed only when select types use it", {
  skip_if_no_writexl()
  p <- make_form(fdf(type = c("text", "select_one_from_file x.csv"),
                     name = c("a", "b"), label = c("A", "B")))
  expect_no_error(read_quiet(p))
  p <- make_form(fdf(type = "select_one l", name = "a", label = "A"))
  expect_error(read_quiet(p), "Sheet `choices` not found.*rows 2")
})

test_that("a select list missing from choices aborts with rows", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "select_one nope", "select_multiple ok",
                     "rank gone", "select_one"),
            name = c("a", "b", "c", "d", "e"), label = c("A", "B", "C", "D", "E"))
  ch <- fdf(list_name = "ok", name = "1", label = "One")
  e <- tryCatch(read_quiet(make_form(sv, ch)),
                error = function(e) conditionMessage(e))
  expect_match(e, "row 3: list `nope`")
  expect_match(e, "row 5: list `gone`")
  expect_match(e, "row 6.*no list name")
  expect_false(grepl("row 4", e))
})

test_that("duplicate names warn with rows", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "text", "text", "note"),
            name = c("a", "b", "a", "a"), label = c("A", "B", "C", "D"))
  expect_warning(read_quiet(make_form(sv)), "`a` on rows 2, 4, 5")
})

test_that("empty names warn on non-structural rows only", {
  skip_if_no_writexl()
  sv <- fdf(type = c("begin_group", "text", "end_group", "note"),
            name = c("g", NA, NA, NA), label = c("G", "A", NA, "N"))
  w <- expect_warning(read_quiet(make_form(sv)), "column name.*empty")
  expect_match(conditionMessage(w), "3, 5")
})

test_that("unknown types give one warning listing rows", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "osm", "foo", NA), name = c("a", "b", "c", "d"),
            label = c("A", "B", "C", "D"))
  w <- NULL
  withCallingHandlers(
    read_quiet(make_form(sv)),
    warning = function(cnd) {
      w <<- c(w, conditionMessage(cnd))
      invokeRestart("muffleWarning")
    }
  )
  unk <- grep("shown like text", w, value = TRUE)
  expect_length(unk, 1)
  expect_match(unk, "row 3: `osm`")
  expect_match(unk, "row 4: `foo`")
  expect_true(any(grepl("empty on 1 row", w)))
  f <- suppressWarnings(read_quiet(make_form(sv)))
  expect_equal(f$survey$kind, c("text", "unknown", "unknown", "unknown"))
})

test_that("unbalanced groups abort when reading", {
  skip_if_no_writexl()
  sv <- fdf(type = c("begin_group", "text"), name = c("g", "a"),
            label = c("G", "A"))
  expect_error(read_quiet(make_form(sv)), "Row 2.*never closed")
})

test_that("required expressions are kept", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "text", "text"), name = c("a", "b", "c"),
            label = c("A", "B", "C"),
            required = c("yes", "${a} = 'x'", "no"))
  f <- read_quiet(make_form(sv))
  expect_equal(f$survey$required, c(TRUE, TRUE, FALSE))
  expect_equal(f$survey$required_expr, c(NA, "${a} = 'x'", NA))
})

test_that("braces in user text do not break messages", {
  skip_if_no_writexl()
  sv <- fdf(type = c("select_one {bad}"), name = "a", label = "A")
  ch <- fdf(list_name = "ok", name = "1", label = "One")
  expect_error(read_quiet(make_form(sv, ch)), "{bad}", fixed = TRUE)
})

test_that("the missing-list error agrees in number", {
  skip_if_no_writexl()
  one <- make_form(fdf(type = "select_one nope", name = "a", label = "A"),
                   fdf(list_name = "l", name = "1", label = "x"))
  expect_error(read_quiet(one), "1 select question uses a choice list")
  two <- make_form(fdf(type = c("select_one nope", "select_one nah"),
                       name = c("a", "b"), label = c("A", "B")),
                   fdf(list_name = "l", name = "1", label = "x"))
  expect_error(read_quiet(two), "2 select questions use a choice list")
})
