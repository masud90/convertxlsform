# Fixture for expression, choice and model tests: a small form with a
# transparent (label-less) group inside a labelled group, and a repeat.
#
# Serials: consent 1, assets 2, hh_size 3, g1 4 (t1 4.1, t2 4.2),
# members 5 (mname 5.1); calc and the last note are unnumbered.
small_form <- function(media_root = NULL, ...) {
  survey <- fdf(
    type = c("select_one yesno", "select_multiple assets", "integer",
             "begin_group", "text", "begin_group", "text", "end_group",
             "end_group", "begin_repeat", "text", "end_repeat", "calculate",
             "note"),
    name = c("consent", "assets", "hh_size", "g1", "t1", "g_hidden", "t2",
             NA, NA, "members", "mname", NA, "calc", "n1"),
    label = c("Consent?", "Assets owned", "Household size", "Group One",
              "Text one", NA, "Text two", NA, NA, "Members", "Member name",
              NA, NA, "Thank you"),
    required = c("yes", rep(NA, 13)),
    relevant = c(NA, NA, "${consent} = 'yes'", NA,
                 "not(selected(${assets}, 'radio')) and (${hh_size} > 2 or ${consent} = 'yes')",
                 "${hh_size} > 1 or ${consent} = 'no'", NA, NA, NA, NA, NA,
                 NA, NA, NA),
    constraint = c(NA, NA, ". >= 0 and . <= 30", rep(NA, 11)),
    constraint_message = c(NA, NA, "Enter 0 to 30", rep(NA, 11)),
    hint = c(NA, NA, "Count everyone", rep(NA, 11)),
    repeat_count = c(rep(NA, 9), "${hh_size}", rep(NA, 4)),
    calculation = c(rep(NA, 12), "1 + 1", NA),
    appearance = c(rep(NA, 5), "field-list", rep(NA, 8))
  )
  choices <- fdf(
    list_name = c("yesno", "yesno", "assets", "assets", "assets"),
    name = c("yes", "no", "radio", "tv", "bike"),
    label = c("Yes", "No", "Radio", "TV", "Bike")
  )
  settings <- fdf(form_title = "Small form", form_id = "small",
                  version = "1")
  path <- make_form(survey, choices, settings)
  read_quiet(path, media_root = media_root, ...)
}

# Default options, quiet.
test_opts <- function(...) {
  .cxf_options(verbose = FALSE, ...)
}

# Block types of a model, in order.
block_types <- function(model) {
  vapply(model$blocks, function(b) b$block, "")
}

# Find a block by name (first match).
block_named <- function(model, name, type = NULL) {
  for (b in model$blocks) {
    if (identical(b$name, name) && (is.null(type) || b$block == type)) {
      return(b)
    }
  }
  NULL
}
