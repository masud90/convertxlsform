# Validation of a parsed XLSForm. ASCII-only source.

#' Escape braces so user text is safe inside cli messages
#' @noRd
.esc <- function(x) {
  x <- gsub("{", "{{", x, fixed = TRUE)
  gsub("}", "}}", x, fixed = TRUE)
}

#' Bullets for a long list of items, capped
#' @noRd
.bullets <- function(x, bullet = "*", max_n = 10L) {
  x <- .esc(x)
  if (length(x) > max_n) {
    extra <- length(x) - max_n
    x <- c(x[seq_len(max_n)], sprintf("... and %d more", extra))
  }
  stats::setNames(x, rep(bullet, length(x)))
}

#' Validate a parsed form
#'
#' Aborts for missing sheets or columns and for select questions whose choice
#' list does not exist. Warns for duplicate or empty names and for unknown
#' types. Every message cites the sheet, column and Excel row.
#'
#' @param form A `cxf_form` (structure not required).
#' @return `form`, invisibly.
#' @noRd
.validate_form <- function(form) {
  sv <- form$survey
  ch <- form$choices
  found <- form$found
  path <- form$path

  # Sheets and columns
  errs <- character(0)
  if (!isTRUE(found$has_survey)) {
    shown <- if (length(form$sheets) > 0L) {
      paste(form$sheets, collapse = ", ")
    } else {
      "none"
    }
    .cxf_abort(c(
      "Sheet {.val survey} not found in {.file {path}}.",
      i = "Sheets found: {.val {shown}}."
    ))
  }
  has_label <- function(h) {
    any(.split_header(h)$field == "label")
  }
  need_sv <- c("type", "name")
  miss <- setdiff(need_sv, found$survey)
  for (m in miss) {
    errs <- c(errs, sprintf("Sheet `survey`, row 1: column `%s` is missing.", m))
  }
  if (!has_label(found$survey)) {
    errs <- c(errs, "Sheet `survey`, row 1: no `label` column (or `label::<language>`).")
  }

  uses_choices <- sv$base_type %in% c("select_one", "select_multiple", "rank")
  if (isTRUE(found$has_choices)) {
    for (m in setdiff(c("list_name", "name"), found$choices)) {
      errs <- c(errs, sprintf("Sheet `choices`, row 1: column `%s` is missing.", m))
    }
    if (!has_label(found$choices)) {
      errs <- c(errs, "Sheet `choices`, row 1: no `label` column (or `label::<language>`).")
    }
  } else if (any(uses_choices)) {
    r <- sv$row[uses_choices]
    errs <- c(errs, sprintf(
      "Sheet `choices` not found, but sheet `survey` has select or rank questions (rows %s).",
      .row_list(r)))
  }
  if (length(errs) > 0L) {
    .cxf_abort(c(
      "The XLSForm {.file {basename(path)}} is missing required sheets or columns.",
      .bullets(errs, "x")
    ))
  }

  # Select questions whose list is missing
  if (any(uses_choices)) {
    lists <- unique(ch$list_name[!is.na(ch$list_name)])
    arg <- sv$type_arg
    bad <- uses_choices & (is.na(arg) | !(arg %in% lists))
    if (any(bad)) {
      msg <- ifelse(
        is.na(sv$type_arg[bad]),
        sprintf("Sheet `survey`, column `type`, row %d: `%s` gives no list name.",
                sv$row[bad], sv$type[bad]),
        sprintf("Sheet `survey`, column `type`, row %d: list `%s` is not in sheet `choices` (column `list_name`).",
                sv$row[bad], sv$type_arg[bad])
      )
      .cxf_abort(c(
        "{sum(bad)} select question{?s} {?uses/use} a choice list that does not exist.",
        .bullets(msg, "x")
      ))
    }
  }

  # Warnings
  blank_type <- is.na(sv$type)
  if (any(blank_type)) {
    rl <- .row_list(sv$row[blank_type])
    nrl <- sum(blank_type)
    .cxf_warn(c(
      "Sheet {.val survey}, column {.field type}: empty on {nrl} row{?s}: {rl}.",
      i = "These rows are treated as unknown types."
    ))
  }
  unk <- sv$kind == "unknown" & !blank_type
  if (any(unk)) {
    msg <- sprintf("row %d: `%s`", sv$row[unk], sv$type[unk])
    .cxf_warn(c(
      "Sheet {.val survey}, column {.field type}: {sum(unk)} unknown type{?s}; shown like text.",
      .bullets(msg)
    ))
  }
  structural_end <- sv$kind %in% c("end_group", "end_repeat")
  no_name <- .is_blank(sv$name) & !structural_end
  if (any(no_name)) {
    rl <- .row_list(sv$row[no_name])
    nrl <- sum(no_name)
    .cxf_warn(c(
      "Sheet {.val survey}, column {.field name}: empty on {nrl} row{?s}: {rl}.",
      i = "Every question, note and group needs a name."
    ))
  }
  has_name <- !.is_blank(sv$name) & !structural_end
  nm <- sv$name[has_name]
  rw <- sv$row[has_name]
  dup <- unique(nm[duplicated(nm)])
  if (length(dup) > 0L) {
    msg <- vapply(dup, function(d) {
      sprintf("`%s` on rows %s", d, .row_list(rw[nm == d]))
    }, "")
    .cxf_warn(c(
      "Sheet {.val survey}, column {.field name}: {length(dup)} duplicate name{?s}.",
      .bullets(msg)
    ))
  }
  invisible(form)
}

#' Compact "3, 7, 12" list of row numbers
#' @noRd
.row_list <- function(x) {
  x <- sort(unique(x))
  if (length(x) > 10L) {
    paste0(paste(x[1:10], collapse = ", "), ", ... (", length(x), " in all)")
  } else {
    paste(x, collapse = ", ")
  }
}
