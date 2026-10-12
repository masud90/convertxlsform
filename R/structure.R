# Group and repeat structure, depth and display numbering.
# ASCII-only source.

#' Add depth, parent and serial to the survey
#'
#' A single pass with an explicit stack. Aborts, citing Excel rows, on
#' unmatched or mismatched `end_*` rows and unclosed `begin_*` rows.
#'
#' @param survey Survey data.frame with `row`, `kind` and `label`; may also
#'   have `appearance`.
#' @param nested_serial_numbers Logical. TRUE numbers groups and nests the
#'   children (2, 2.1, 2.2.1). FALSE numbers questions 1..n and leaves groups
#'   unnumbered.
#' @return `survey` with `depth` (int), `parent` (int, NA at top level) and
#'   `serial` (chr, NA when unnumbered).
#' @noRd
.build_structure <- function(survey, nested_serial_numbers = TRUE) {
  n <- nrow(survey)
  kind <- as.character(survey$kind)
  rows <- survey$row
  has_label <- !.is_blank(survey$label)

  depth <- integer(n)
  parent <- rep(NA_integer_, n)
  serial <- rep(NA_character_, n)

  # Stack of open begin_* rows.
  st_idx <- integer(n + 1L)   # survey index of the begin row
  st_nf <- integer(n + 1L)    # numbering frame used by the children
  top <- 0L
  # Numbering frames: prefix and counter. Frame 1 is the top level.
  prefix <- character(n + 2L)
  counter <- integer(n + 2L)
  nframes <- 1L
  global <- 0L

  errors <- character(0)
  unnumbered <- c("note", "hidden", "end_group", "end_repeat")

  for (i in seq_len(n)) {
    k <- kind[i]
    cur_nf <- if (top > 0L) st_nf[top] else 1L
    depth[i] <- top
    if (top > 0L) {
      parent[i] <- st_idx[top]
    }

    if (k %in% c("begin_group", "begin_repeat")) {
      transparent <- k == "begin_group" && !has_label[i]
      new_nf <- cur_nf
      if (nested_serial_numbers && !transparent) {
        counter[cur_nf] <- counter[cur_nf] + 1L
        s <- if (cur_nf == 1L) {
          as.character(counter[cur_nf])
        } else {
          paste0(prefix[cur_nf], ".", counter[cur_nf])
        }
        serial[i] <- s
        nframes <- nframes + 1L
        prefix[nframes] <- s
        counter[nframes] <- 0L
        new_nf <- nframes
      }
      top <- top + 1L
      st_idx[top] <- i
      st_nf[top] <- new_nf
    } else if (k %in% c("end_group", "end_repeat")) {
      want <- if (k == "end_group") "begin_group" else "begin_repeat"
      if (top == 0L) {
        errors <- c(errors, sprintf(
          "Row %d: `%s` has no matching `%s`.", rows[i],
          if (k == "end_group") "end_group" else "end_repeat", want))
      } else {
        b <- st_idx[top]
        if (kind[b] != want) {
          errors <- c(errors, sprintf(
            "Row %d: `%s` closes `%s` opened at row %d.", rows[i], k,
            kind[b], rows[b]))
        }
        top <- top - 1L
        # The end row sits at the level of its begin row.
        depth[i] <- top
        parent[i] <- if (top > 0L) st_idx[top] else NA_integer_
      }
    } else if (!(k %in% unnumbered)) {
      if (nested_serial_numbers) {
        counter[cur_nf] <- counter[cur_nf] + 1L
        serial[i] <- if (cur_nf == 1L) {
          as.character(counter[cur_nf])
        } else {
          paste0(prefix[cur_nf], ".", counter[cur_nf])
        }
      } else {
        global <- global + 1L
        serial[i] <- as.character(global)
      }
    }
  }

  if (top > 0L) {
    open <- st_idx[seq_len(top)]
    errors <- c(errors, sprintf(
      "Row %d: `%s` is never closed.", rows[open], kind[open]))
  }
  if (length(errors) > 0L) {
    errors <- .esc(errors)
    .cxf_abort(c(
      "Unbalanced groups or repeats in sheet {.val survey}.",
      stats::setNames(errors, rep("x", length(errors)))
    ))
  }

  survey$depth <- depth
  survey$parent <- parent
  survey$serial <- serial
  survey
}
