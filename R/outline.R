# The outline: a flat view of the document model. ASCII-only source.

#' Outline of an XLSForm
#'
#' Parses an XLSForm and returns one row for each item that would appear on
#' the paper form: sections, questions, notes and the end of repeat groups. It
#' shows the question numbers, the answer type, the number of choices, and the
#' skip logic and constraints in plain language, so a form can be checked
#' before it is converted.
#'
#' @param xlsform Path to an XLSForm (`.xlsx` or `.xls`).
#' @param lang Language to use: a code (`"fr"`), a name (`"French"`) or the
#'   full header suffix (`"French (fr)"`). `NULL` uses the `default_language`
#'   of the settings sheet, or the first language found. See
#'   [xlsform_languages()].
#' @param ... Further options: `nested_serial_numbers`, `include_relevants`,
#'   `include_constraints`, `include_hints`, `include_hidden`, `media_root`,
#'   `max_choices` (lists longer than this get a code box layout), `strings`,
#'   `show_variable_names`, `title` and the other conversion options. Page and font
#'   options are accepted and have no effect here.
#'
#' @return A data.frame with one row per block and columns `row` (Excel row of
#'   the form), `serial` (question number, NA when unnumbered), `depth`
#'   (number of enclosing groups), `block` (`"section"`, `"section_end"`,
#'   `"note"` or `"question"`), `kind` (the question kind, for example
#'   `"select_one"`, or `"begin_group"`, `"begin_repeat"`, `"end_repeat"`,
#'   `"note"`), `name`, `label`, `required`, `n_choices` (all choices of a
#'   select or rank question, not counting an "other" row), `relevant_text`,
#'   `constraint_text` (the constraint message when the form has one, else the
#'   humanized expression) and `layout` (for select and rank questions:
#'   `"list"`, `"columns"` or `"code_box"`, from the number of choices and
#'   `max_choices`; NA otherwise). A label may contain line breaks (`\n`).
#'
#' @examples
#' path <- system.file("extdata", "minimal", "minimal.xlsx",
#'                     package = "convertxlsform")
#' out <- xlsform_outline(path)
#' out[, c("serial", "block", "kind", "name", "relevant_text")]
#'
#' \donttest{
#' # A multilingual form: another language, with flat question numbers
#' path_ml <- system.file("extdata", "multilang", "multilang.xlsx",
#'                        package = "convertxlsform")
#' out <- xlsform_outline(path_ml, lang = "bn", nested_serial_numbers = FALSE)
#' head(out[, c("serial", "name")], 5)
#' }
#' @export
xlsform_outline <- function(xlsform, lang = NULL, ...) {
  call <- rlang::current_call()
  .with_call(call, {
    args <- list(...)
    if (is.null(args$verbose)) {
      args$verbose <- FALSE
    }
    opts <- do.call(.cxf_options, c(list(lang = lang), args))
    form <- .read_xlsform(xlsform, lang = opts$lang,
                          media_root = opts$media_root,
                          verbose = opts$verbose)
    .outline_from_model(.build_model(form, opts))
  })
}

#' Flatten a cxf_model into the outline data.frame
#' @noRd
.outline_from_model <- function(model) {
  blocks <- model$blocks
  get <- function(field, type, na) {
    vapply(blocks, function(b) {
      v <- b[[field]]
      if (is.null(v) || length(v) != 1L) na else v
    }, type, USE.NAMES = FALSE)
  }
  kind <- vapply(blocks, function(b) {
    switch(
      b$block,
      section = if (isTRUE(b$is_repeat)) "begin_repeat" else "begin_group",
      section_end = "end_repeat",
      note = "note",
      b$kind
    )
  }, "", USE.NAMES = FALSE)
  n_choices <- vapply(blocks, function(b) {
    n <- if (identical(b$block, "question")) b$answer$n_total else NULL
    if (is.null(n)) NA_integer_ else as.integer(n)
  }, 1L, USE.NAMES = FALSE)
  layout <- vapply(blocks, function(b) {
    l <- if (identical(b$block, "question")) b$answer$layout else NULL
    if (is.null(l)) NA_character_ else l
  }, "", USE.NAMES = FALSE)
  data.frame(
    row = get("row", 1L, NA_integer_),
    serial = get("serial", "", NA_character_),
    depth = get("depth", 1L, NA_integer_),
    block = get("block", "", NA_character_),
    kind = kind,
    name = get("name", "", NA_character_),
    label = get("label", "", NA_character_),
    required = get("required", NA, NA),
    n_choices = n_choices,
    relevant_text = get("relevant_text", "", NA_character_),
    constraint_text = get("constraint_text", "", NA_character_),
    layout = layout,
    stringsAsFactors = FALSE
  )
}
