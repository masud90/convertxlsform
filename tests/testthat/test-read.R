test_that("sheet and column names are matched case-insensitively", {
  skip_if_no_writexl()
  p <- make_form(
    fdf(` Type ` = c("text", "integer"), NAME = c("a", "b"),
        Label = c("A", "B"), REQUIRED = c("yes", NA)),
    choices = fdf(List_Name = "l", Name = "1", Label = "One"), survey_name = "Survey ", choices_name = "CHOICES",
    settings = fdf(Form_Title = "T", FORM_ID = "id1"),
    settings_name = " Settings"
  )
  f <- read_quiet(p)
  expect_s3_class(f, "cxf_form")
  expect_equal(f$survey$name, c("a", "b"))
  expect_equal(f$survey$label, c("A", "B"))
  expect_equal(f$survey$required, c(TRUE, FALSE))
  expect_equal(f$settings$form_title, "T")
  expect_equal(f$settings$form_id, "id1")
  expect_true(is.na(f$settings$version))
  expect_true(is.na(f$settings$default_language))
})

test_that("text columns keep leading zeros and blank rows are dropped", {
  skip_if_no_writexl()
  sv <- fdf(type = c("select_one l", NA, "text", NA, "integer"),
            name = c("q1", NA, "q2", NA, "q3"),
            label = c("Q1", NA, "Q2", NA, "Q3"),
            default = c("01", NA, NA, NA, "007"))
  ch <- fdf(list_name = c("l", "l"), name = c("01", "02"),
            label = c("One", "Two"))
  f <- read_quiet(make_form(sv, ch))
  expect_equal(f$survey$row, c(2L, 4L, 6L))
  expect_equal(f$survey$default, c("01", NA, "007"))
  expect_equal(f$choices$name, c("01", "02"))
  expect_equal(f$choices$row, 2:3)
})

test_that("leading blank rows are not skipped from Excel numbering", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "text"), name = c("a", "b"), label = c("A", "B"))
  p <- make_form(sv)
  f <- read_quiet(p)
  expect_equal(f$survey$row, c(2L, 3L))
})

test_that("elements of the cxf_form follow the spec", {
  skip_if_no_writexl()
  sv <- fdf(type = c("text", "select_one yn or_other"), name = c("a", "b"),
            label = c("A", "B"), parameters = c("x=1", NA))
  ch <- fdf(list_name = c("yn", "yn"), name = c("1", "2"),
            label = c("Yes", "No"), region = c("n", "s"))
  f <- read_quiet(make_form(sv, ch), media_root = tempdir())
  expect_named(f, c("path", "media_root", "settings", "languages", "lang",
                    "survey", "choices", "sheets", "found"))
  expect_equal(f$media_root, normalizePath(tempdir(), winslash = "/"))
  expect_equal(nrow(f$languages), 0)
  expect_equal(f$lang$mode, "single")
  expect_true(is.na(f$lang$suffix))
  expect_true(all(c("row", "type", "base_type", "type_arg", "or_other", "kind",
                    "name", "label", "hint", "guidance_hint",
                    "required_message", "constraint_message", "image", "audio",
                    "video", "required", "required_expr", "relevant",
                    "constraint", "choice_filter", "repeat_count",
                    "calculation", "default", "appearance", "parameters",
                    "read_only", "params", "depth", "parent", "serial")
                  %in% names(f$survey)))
  expect_equal(f$survey$base_type, c("text", "select_one"))
  expect_equal(f$survey$type_arg, c(NA, "yn"))
  expect_equal(f$survey$or_other, c(FALSE, TRUE))
  expect_equal(f$survey$params[[1]], list(x = 1))
  expect_length(f$survey$params[[2]], 0)
  expect_equal(f$survey$serial, c("1", "2"))
  expect_equal(names(f$choices)[1:5],
               c("list_name", "name", "label", "image", "row"))
  expect_equal(f$choices$region, c("n", "s"))
  expect_equal(dirname(f$path), dirname(normalizePath(f$path)))
})

test_that("type whitespace and capitalization are normalized", {
  skip_if_no_writexl()
  sv <- fdf(type = c("  begin   group ", "dateTime", "end group"),
            name = c("g", "d", NA), label = c("G", "D", NA))
  f <- read_quiet(make_form(sv))
  expect_equal(f$survey$type[1], "begin group")
  expect_equal(f$survey$base_type, c("begin_group", "datetime", "end_group"))
  expect_equal(f$survey$kind, c("begin_group", "datetime", "end_group"))
})

test_that("bad files give clean errors naming the file", {
  expect_error(.read_xlsform(file.path(tempdir(), "nope.xlsx")), "not found")
  txt <- tempfile(fileext = ".xlsx")
  writeLines("not an excel file", txt)
  expect_error(.read_xlsform(txt), "Cannot read")
  expect_error(.read_xlsform(txt), basename(txt), fixed = TRUE)
  csv <- tempfile(fileext = ".csv")
  writeLines("a,b", csv)
  expect_error(.read_xlsform(csv), "not an Excel file")
  expect_error(.read_xlsform(c("a", "b")), "single file path")
})

test_that("multiple settings and extra sheets do not disturb reading", {
  skip_if_no_writexl()
  sv <- fdf(type = "text", name = "a", label = "A")
  f <- read_quiet(make_form(sv, other = fdf(x = "1")))
  expect_equal(f$sheets, c("survey", "other"))
})

test_that("verbose reading announces the file", {
  skip_if_no_writexl()
  p <- make_form(fdf(type = "text", name = "a", label = "A"))
  expect_message(.read_xlsform(p, verbose = TRUE), "Reading")
  expect_no_message(.read_xlsform(p, verbose = FALSE))
})
