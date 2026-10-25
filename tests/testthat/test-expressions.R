test_that("references show the question number and name", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x, field = "relevant") .humanize_expr(x, f, field)
  expect_equal(h("${hh_size} > 2"), "Q3 (hh_size) > 2")
  expect_equal(h("${ hh_size } > 2"), "Q3 (hh_size) > 2")
  # unnumbered (calculate) and unknown references show the bare name
  expect_equal(h("${calc} > 2"), "calc > 2")
  expect_equal(h("${nothing} > 2"), "nothing > 2")
  # a nested serial
  expect_equal(h("${t2} = 'x'"), "Q4.2 (t2) = \"x\"")
})

test_that("comparison operators follow the spec", {
  skip_if_no_writexl()
  f <- small_form()
  cases <- c(
    "${hh_size} > 2" = "Q3 (hh_size) > 2",
    "${hh_size} < 2" = "Q3 (hh_size) < 2",
    "${hh_size} >= 2" = "Q3 (hh_size) >= 2",
    "${hh_size} <= 2" = "Q3 (hh_size) <= 2",
    "${hh_size} = 2" = "Q3 (hh_size) = 2",
    "${hh_size} != 2" = "Q3 (hh_size) is not 2",
    "${hh_size}>2" = "Q3 (hh_size) > 2",
    "${hh_size}>=2.5" = "Q3 (hh_size) >= 2.5",
    "${hh_size} + 1 > 3" = "Q3 (hh_size) + 1 > 3",
    "${hh_size} * 2 = 4" = "Q3 (hh_size) * 2 = 4",
    "${hh_size} div 2 = 1" = "Q3 (hh_size) / 2 = 1",
    "${hh_size} mod 2 = 0" = "Q3 (hh_size) mod 2 = 0",
    "-${hh_size} < 0" = "-Q3 (hh_size) < 0",
    "${hh_size} - 1 > 0" = "Q3 (hh_size) - 1 > 0"
  )
  for (i in seq_along(cases)) {
    expect_equal(.humanize_expr(names(cases)[i], f, "relevant"), unname(cases[i]),
                 info = names(cases)[i])
  }
})

test_that("select_one comparisons use the choice label", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x) .humanize_expr(x, f, "relevant")
  expect_equal(h("${consent} = 'yes'"), "Q1 (consent) is \"Yes\"")
  expect_equal(h("${consent} = \"yes\""), "Q1 (consent) is \"Yes\"")
  expect_equal(h("${consent}='no'"), "Q1 (consent) is \"No\"")
  expect_equal(h("${consent} != 'yes'"), "Q1 (consent) is not \"Yes\"")
  expect_equal(h("'yes' = ${consent}"), "Q1 (consent) is \"Yes\"")
  # a value that is not in the list is kept as written
  expect_equal(h("${consent} = 'maybe'"), "Q1 (consent) is \"maybe\"")
  # select_multiple compared with = is not relabelled
  expect_equal(h("${assets} = 'radio'"), "Q2 (assets) = \"radio\"")
})

test_that("selected, not(selected) and count-selected", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x) .humanize_expr(x, f, "relevant")
  expect_equal(h("selected(${assets}, 'radio')"),
               "Q2 (assets) includes \"Radio\"")
  expect_equal(h("selected(${assets},\"tv\")"), "Q2 (assets) includes \"TV\"")
  expect_equal(h("selected(${assets}, 'zzz')"), "Q2 (assets) includes \"zzz\"")
  expect_equal(h("not(selected(${assets}, 'radio'))"),
               "Q2 (assets) does not include \"Radio\"")
  expect_equal(h("not( selected( ${assets} , 'radio' ) )"),
               "Q2 (assets) does not include \"Radio\"")
  expect_equal(h("count-selected(${assets}) > 1"),
               "number selected in Q2 (assets) > 1")
  expect_equal(h("count-selected(${assets}) = 0"),
               "number selected in Q2 (assets) = 0")
  # the second argument can be another question
  expect_equal(h("selected(${assets}, ${consent})"),
               "Q2 (assets) includes Q1 (consent)")
  # select_one used with selected(): still labelled
  expect_equal(h("selected(${consent}, 'yes')"),
               "Q1 (consent) includes \"Yes\"")
})

test_that("not() of other expressions", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x) .humanize_expr(x, f, "relevant")
  expect_equal(h("not(${consent} = 'yes')"), "Q1 (consent) is not \"Yes\"")
  expect_equal(h("not(${consent} != 'yes')"), "Q1 (consent) is \"Yes\"")
  expect_equal(h("not(${hh_size} > 2)"), "not (Q3 (hh_size) > 2)")
  expect_equal(h("not(${hh_size} > 2 and ${hh_size} < 9)"),
               "not (Q3 (hh_size) > 2 and Q3 (hh_size) < 9)")
})

test_that("nested logic keeps the structure", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x) .humanize_expr(x, f, "relevant")
  expect_equal(
    h("not(selected(${assets}, 'x')) and (${hh_size} > 2 or ${consent} = 'yes')"),
    "Q2 (assets) does not include \"x\" and (Q3 (hh_size) > 2 or Q1 (consent) is \"Yes\")"
  )
  expect_equal(
    h("(${hh_size} > 2 and (${consent} = 'yes' or selected(${assets}, 'tv'))) or ${hh_size} = 0"),
    "(Q3 (hh_size) > 2 and (Q1 (consent) is \"Yes\" or Q2 (assets) includes \"TV\")) or Q3 (hh_size) = 0"
  )
  # operators are case-sensitive, as in XPath: left as written
  expect_equal(h("${hh_size} > 2 AND ${hh_size} < 4"),
               "Q3 (hh_size) > 2 AND Q3 (hh_size) < 4")
})

test_that("empty strings and today()", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x, field = "relevant") .humanize_expr(x, f, field)
  expect_equal(h("${consent} != ''"), "Q1 (consent) is not empty")
  expect_equal(h("${consent} = ''"), "Q1 (consent) is empty")
  expect_equal(h("${consent} != \"\""), "Q1 (consent) is not empty")
  expect_equal(h(". = today()", "constraint"), "the answer = today")
  expect_equal(h(". <= today()", "constraint"), "the answer <= today")
  expect_equal(h("${hh_size} > 1 and true()"), "Q3 (hh_size) > 1 and true")
})

test_that("string quoting handles both quote styles", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x) .humanize_expr(x, f, "relevant")
  expect_equal(h("${t2} = 'a b'"), "Q4.2 (t2) = \"a b\"")
  expect_equal(h("${t2} = \"a b\""), "Q4.2 (t2) = \"a b\"")
  # a double quote inside single quotes is shown with single quotes
  expect_equal(h("${t2} = 'say \"hi\"'"), "Q4.2 (t2) = 'say \"hi\"'")
  # a single quote inside double quotes
  expect_equal(h("${t2} = \"it's\""), "Q4.2 (t2) = \"it's\"")
  # operators inside strings are not operators
  expect_equal(h("${t2} = 'a and b'"), "Q4.2 (t2) = \"a and b\"")
})

test_that("a lone dot depends on the field", {
  skip_if_no_writexl()
  f <- small_form()
  expect_equal(.humanize_expr(". >= 0", f, "constraint"),
               "the answer >= 0")
  expect_equal(.humanize_expr(". >= 0 and . <= 30", f, "constraint"),
               "the answer >= 0 and the answer <= 30")
  expect_equal(.humanize_expr(".>=0", f, "constraint"), "the answer >= 0")
  expect_equal(.humanize_expr(". >= 0", f, "relevant"), "this answer >= 0")
  expect_equal(.humanize_expr(". >= 0", f, "choice_filter"),
               "this answer >= 0")
  expect_equal(.humanize_expr(". >= 0", f, "repeat_count"),
               "this answer >= 0")
  # the dot takes the choice labels of its own question when the row is given
  i <- which(f$survey$name == "assets")
  expect_equal(.humanize_expr("selected(., 'tv')", f, "constraint", row = i),
               "the answer includes \"TV\"")
  expect_equal(.humanize_expr("selected(., 'tv')", f, "constraint"),
               "the answer includes \"tv\"")
  j <- which(f$survey$name == "consent")
  expect_equal(.humanize_expr(". = 'yes'", f, "constraint", row = j),
               "the answer is \"Yes\"")
})

test_that("repeat_count and choice_filter", {
  skip_if_no_writexl()
  f <- small_form()
  expect_equal(.humanize_expr("${hh_size}", f, "repeat_count"),
               "Q3 (hh_size)")
  expect_equal(.humanize_expr("3", f, "repeat_count"), "3")
  expect_equal(.humanize_expr("district = ${consent}", f, "choice_filter"),
               "district = Q1 (consent)")
})

test_that("unknown functions and odd syntax fall back to substitution only", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x, field = "relevant") .humanize_expr(x, f, field)
  expect_equal(h("substr(., 0, 1) = 'a'", "constraint"),
               "substr(., 0, 1) = 'a'")
  expect_equal(h("format-date(${t2}, '%Y') = '2020' and ${consent} = 'yes'"),
               "format-date(Q4.2 (t2), '%Y') = '2020' and Q1 (consent) = 'yes'")
  expect_equal(h("/data/x = 1"), "/data/x = 1")
  expect_equal(h("(${hh_size} > 2"), "(Q3 (hh_size) > 2")
  expect_equal(h("${hh_size} >"), "Q3 (hh_size) >")
  expect_equal(h("selected(${assets})"), "selected(Q2 (assets))")
  expect_equal(h("count-selected(${assets}, 1)"),
               "count-selected(Q2 (assets), 1)")
  expect_equal(h("${hh_size} > 2 or ?"), "Q3 (hh_size) > 2 or ?")
  # a known function with the wrong number of arguments
  expect_equal(h("string-length(${t2}, 2) > 3"),
               "string-length(Q4.2 (t2), 2) > 3")
  # a path inside position() or count()
  expect_equal(h("position(../x) = 1"), "position(../x) = 1")
})

test_that("NA and blank input", {
  skip_if_no_writexl()
  f <- small_form()
  expect_true(is.na(.humanize_expr(NA, f, "relevant")))
  expect_true(is.na(.humanize_expr(NA_character_, f, "constraint")))
  expect_true(is.na(.humanize_expr("", f, "relevant")))
  expect_true(is.na(.humanize_expr("   ", f, "relevant")))
  expect_error(.humanize_expr("1", f, "label"), "field")
})

test_that(".humanize_exprs is vectorized", {
  skip_if_no_writexl()
  f <- small_form()
  x <- c("${hh_size} > 2", NA, "selected(${assets}, 'tv')", "")
  expect_equal(
    .humanize_exprs(x, f, "relevant"),
    c("Q3 (hh_size) > 2", NA, "Q2 (assets) includes \"TV\"", NA)
  )
  expect_equal(.humanize_exprs(character(0), f, "relevant"), character(0))
  i <- which(f$survey$name == "assets")
  expect_equal(
    .humanize_exprs(c("selected(., 'tv')", "selected(., 'bike')"), f,
                    "constraint", rows = c(i, i)),
    c("the answer includes \"TV\"", "the answer includes \"Bike\"")
  )
})

test_that("choice labels follow the chosen language", {
  p <- system.file("extdata", "multilang", "multilang.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  en <- read_quiet(p, lang = "en")
  bn <- read_quiet(p, lang = "bn")
  x <- "${dwelling_ownership}='1'"
  expect_equal(.humanize_expr(x, en, "relevant"),
               "Q3.5 (dwelling_ownership) is \"Yes\"")
  expect_match(.humanize_expr(x, bn, "relevant"), "^Q3.5 \\(dwelling_ownership\\) is ")
  expect_false(identical(.humanize_expr(x, en, "relevant"),
                         .humanize_expr(x, bn, "relevant")))
  # the multilang example's constraint uses if()
  con <- en$survey$constraint[which(en$survey$name == "asset_owned")]
  out <- .humanize_expr(con, en, "constraint")
  expect_match(out, "^if Q4.1 \\(asset_owned\\) includes \"Refused to respond\"")
  expect_match(out, ", otherwise true$")
})

test_that("simple numeric ranges are described", {
  r <- .simple_range_text
  expect_equal(r(". >= 0 and . <= 30"), "from 0 to 30")
  expect_equal(r(".>=0 and .<=30"), "from 0 to 30")
  expect_equal(r(". <= 30 and . >= 0"), "from 0 to 30")
  expect_equal(r(". > 0 and . <= 30"), "more than 0 and at most 30")
  expect_equal(r(". > 1 and . < 2"), "more than 1 and less than 2")
  expect_equal(r(". >= 0"), "at least 0")
  expect_equal(r(". < 5"), "less than 5")
  fr <- .cxf_strings("fr")
  expect_equal(.simple_range_text(". >= 0 and . <= 30", fr),
               "de 0 \u00e0 30")
  expect_equal(.simple_range_text(". > 0 and . <= 30", fr),
               "plus de 0 et au plus 30")
  expect_equal(.simple_range_text(". >= 2", fr), "au moins 2")
  expect_equal(.simple_range_text(". < 5", fr), "moins de 5")
  expect_true(is.na(r(". >= 1900 or . <= 2025")))
  expect_true(is.na(r(". >= 0 and . >= 1")))
  expect_true(is.na(r("regex(., '^[0-9]+$')")))
  expect_true(is.na(r(". >= ${hh_size}")))
  expect_true(is.na(r("(. > 1")))
  expect_true(is.na(r(NA)))
  expect_true(is.na(r("")))
})

test_that("show_variable_names switches between Q3 and Q3 (name)", {
  skip_if_no_writexl()
  f <- small_form()
  on <- function(x, field = "relevant") {
    .humanize_expr(x, f, field, show_variable_names = TRUE)
  }
  off <- function(x, field = "relevant") {
    .humanize_expr(x, f, field, show_variable_names = FALSE)
  }
  expect_equal(on("${hh_size} > 2"), "Q3 (hh_size) > 2")
  expect_equal(off("${hh_size} > 2"), "Q3 > 2")
  expect_equal(off("${consent} = 'yes'"), "Q1 is \"Yes\"")
  expect_equal(off("not(selected(${assets}, 'radio'))"),
               "Q2 does not include \"Radio\"")
  expect_equal(off("count-selected(${assets}) > 1"),
               "number selected in Q2 > 1")
  expect_equal(off("${t2} = 'x'"), "Q4.2 = \"x\"")
  # unnumbered and unknown references keep the name either way
  expect_equal(off("${calc} > 2"), "calc > 2")
  # the verbatim fallback is substituted the same way
  expect_equal(off("substr(${t2}, 0, 1)"), "substr(Q4.2, 0, 1)")
  expect_equal(on("substr(${t2}, 0, 1)"), "substr(Q4.2 (t2), 0, 1)")
  v <- .humanize_exprs(c("${hh_size} > 2", "${hh_size}"), f, "relevant",
                       show_variable_names = FALSE)
  expect_equal(v, c("Q3 > 2", "Q3"))
})

test_that("with_fallback flags expressions shown as written", {
  skip_if_no_writexl()
  f <- small_form()
  r <- .humanize_exprs(c("${hh_size} > 2", "substr(., 0, 1) = 'a'", NA), f,
                       "constraint", with_fallback = TRUE)
  expect_named(r, c("text", "fallback"))
  expect_equal(r$fallback, c(FALSE, TRUE, FALSE))
  expect_equal(r$text[1], "Q3 (hh_size) > 2")
})

test_that("humanized text is built from the French string table", {
  skip_if_no_writexl()
  f <- small_form()
  fr <- .cxf_strings("fr")
  h <- function(x, field = "relevant", ...) {
    .humanize_expr(x, f, field, strings = fr, show_variable_names = FALSE, ...)
  }
  expect_equal(h("${consent} = 'yes'"), "Q1 est \"Yes\"")
  expect_equal(h("${consent} != 'yes'"), "Q1 n'est pas \"Yes\"")
  expect_equal(h("selected(${assets}, 'tv')"), "Q2 comprend \"TV\"")
  expect_equal(h("not(selected(${assets}, 'tv'))"),
               "Q2 ne comprend pas \"TV\"")
  expect_equal(h("${consent} = ''"), "Q1 est vide")
  expect_equal(h("${consent} != ''"), "Q1 n'est pas vide")
  expect_equal(h("${hh_size} > 1 and ${hh_size} < 9 or true()"),
               "Q3 > 1 et Q3 < 9 ou vrai")
  expect_equal(h(". <= today()", "constraint"),
               "la r\u00e9ponse <= aujourd'hui")
  expect_equal(h(". > 0", "relevant"), "cette r\u00e9ponse > 0")
  expect_equal(h("count-selected(${assets}) > 1"),
               paste("nombre de choix s\u00e9lectionn\u00e9s dans Q2 > 1"))
  expect_equal(h("not(${hh_size} > 2)"), "non (Q3 > 2)")
  # overrides apply
  ov <- .cxf_strings("fr", list(includes = "contient"))
  expect_equal(.humanize_expr("selected(${assets}, 'tv')", f, strings = ov),
               "Q2 (assets) contient \"TV\"")
})

test_that("the form language selects the strings by default", {
  p <- system.file("extdata", "all_types", "all_types.xlsx",
                   package = "convertxlsform")
  skip_if(!nzchar(p))
  fr <- read_quiet(p, lang = "fr")
  en <- read_quiet(p, lang = "en")
  x <- "${consent} = ''"
  expect_match(.humanize_expr(x, fr, "relevant"), "est vide$")
  expect_match(.humanize_expr(x, en, "relevant"), "is empty$")
})

test_that("common XLSForm functions are written in plain language", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x, field = "relevant") .humanize_expr(x, f, field)
  # constraints on text
  expect_equal(h("string-length(.) <= 10", "constraint"),
               "number of characters in the answer <= 10")
  expect_equal(h("regex(., '^[0-9]{10}$')", "constraint"),
               "the answer matches the pattern \"^[0-9]{10}$\"")
  expect_equal(h("regex(${t2}, ${t1})"), "Q4.2 (t2) matches the pattern Q4.1 (t1)")
  # choices
  expect_equal(h("selected-at(${assets}, 0) = 'tv'"),
               "selected choice no. 1 in Q2 (assets) = \"tv\"")
  expect_equal(h("jr:choice-name(${consent}, '${consent}')"),
               "label of the choice in Q1 (consent)")
  # repeats
  expect_equal(h("count(${members}) > 2"),
               "number of entries in Q5 (members) > 2")
  expect_equal(h("position(..) = 1"), "number of the current entry = 1")
  expect_equal(h("position() > 1"), "number of the current entry > 1")
  expect_equal(h("indexed-repeat(${mname}, ${members}, 1) = 'Ann'"),
               "Q5.1 (mname) in entry 1 of Q5 (members) = \"Ann\"")
  # if() and coalesce(), bracketed when they are compared
  expect_equal(h("if(${hh_size} > 2, 'big', 'small')"),
               "if Q3 (hh_size) > 2 then \"big\", otherwise \"small\"")
  expect_equal(h("if(${hh_size} > 2, 1, 0) = 1"),
               "(if Q3 (hh_size) > 2 then 1, otherwise 0) = 1")
  expect_equal(h("coalesce(${t1}, ${t2}) = 'x'"),
               "(Q4.1 (t1), or Q4.2 (t2) if empty) = \"x\"")
  # numbers and dates
  expect_equal(h("int((today() - ${hh_size}) div 365.25) >= 18"),
               "years since Q3 (hh_size) >= 18")
  expect_equal(h("int((today() - ${hh_size}) div 365) < 5"),
               "years since Q3 (hh_size) < 5")
  expect_equal(h("int(${hh_size} div 2) = 1"),
               "whole part of Q3 (hh_size) / 2 = 1")
  expect_equal(h("round(${hh_size}) = 2"), "Q3 (hh_size) rounded = 2")
  expect_equal(h("round(${hh_size}, 1) > 2"),
               "Q3 (hh_size) rounded to 1 decimal places > 2")
  expect_equal(h("number(${t1}) > 2"), "Q4.1 (t1) > 2")
  expect_equal(h(". >= date('2020-01-01')", "constraint"),
               "the answer >= \"2020-01-01\"")
  expect_equal(h("decimal-date-time(.) < decimal-date-time(today())",
                 "constraint"), "the answer < today")
  # values computed on the device
  expect_equal(h("pulldata('a', 'b', 'c', ${consent}) = 'x'"),
               "a calculated value = \"x\"")
  expect_equal(h("once(uuid())"), "a calculated value")
})

test_that("unusual uses of the new functions fall back to the raw text", {
  skip_if_no_writexl()
  f <- small_form()
  h <- function(x, field = "relevant") .humanize_expr(x, f, field)
  # index that is not a whole number literal
  expect_equal(h("selected-at(${assets}, ${hh_size}) = 'tv'"),
               "selected-at(Q2 (assets), Q3 (hh_size)) = 'tv'")
  # the question is not given as '${name}'
  expect_equal(h("jr:choice-name(${consent}, '/data/consent')"),
               "jr:choice-name(Q1 (consent), '/data/consent')")
  # wrong number of arguments
  expect_equal(h("if(${hh_size} > 2, 1)"), "if(Q3 (hh_size) > 2, 1)")
  expect_equal(h("indexed-repeat(${mname}, ${members})"),
               "indexed-repeat(Q5.1 (mname), Q5 (members))")
  r <- .humanize_exprs("regex(., 'a', 'b')", f, "constraint",
                       with_fallback = TRUE)
  expect_true(r$fallback)
})

test_that("the function wording comes from the string table", {
  skip_if_no_writexl()
  f <- small_form()
  fr <- .cxf_strings("fr")
  h <- function(x, field = "relevant") {
    .humanize_expr(x, f, field, strings = fr, show_variable_names = FALSE)
  }
  expect_equal(h("string-length(.) <= 10", "constraint"),
               "nombre de caract\u00e8res de la r\u00e9ponse <= 10")
  expect_equal(h("if(${hh_size} > 2, 1, 0) = 1"),
               "(si Q3 > 2 alors 1, sinon 0) = 1")
  expect_equal(h("int((today() - ${hh_size}) div 365.25) >= 18"),
               "ann\u00e9es \u00e9coul\u00e9es depuis Q3 >= 18")
  # an override reaches the template
  ov <- .cxf_strings("en", list(years_since = "age from %s"))
  expect_equal(.humanize_expr("int((today() - ${hh_size}) div 365.25)", f,
                              strings = ov, show_variable_names = FALSE),
               "age from Q3")
})
