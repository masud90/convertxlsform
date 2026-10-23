# Forward skip instructions ("Go to Q5") for single-choice questions.
# ASCII-only source.
#
# For each choice of a select_one question S, the relevance of the rows that
# follow S is evaluated with S set to that choice. When the choice hides one
# or more rows, the choice gets the number of the next row that is still
# asked. Relevance that depends on S together with other answers is only used
# when it is decided by S alone (three-valued logic); otherwise no
# instruction is printed for that choice.

#' Names referenced as `${name}` in a parsed expression
#' @noRd
.sk_refs <- function(node) {
  if (is.null(node) || !is.list(node)) {
    return(character(0))
  }
  if (identical(node$t, "ref")) {
    return(node$name)
  }
  kids <- c(list(node$x, node$l, node$r), node$args %||% list())
  unique(unlist(lapply(kids, .sk_refs), use.names = FALSE))
}

#' Value of a literal node as a string (NULL for other nodes)
#' @noRd
.sk_lit <- function(node) {
  switch(node$t, str = node$value, num = node$text, NULL)
}

#' Is the node a reference to `s`?
#' @noRd
.sk_is_s <- function(node, s) {
  identical(node$t, "ref") && identical(node$name, s)
}

#' Evaluate a parsed relevance with question `s` answered `v`
#'
#' @return TRUE, FALSE or NA (cannot be decided from `s` alone).
#' @noRd
.sk_eval <- function(node, s, v) {
  ev <- function(x) .sk_eval(x, s, v)
  switch(
    node$t,
    paren = ev(node$x),
    bin = {
      op <- node$op
      if (op == "and") {
        a <- ev(node$l)
        if (isFALSE(a)) return(FALSE)
        b <- ev(node$r)
        if (isFALSE(b)) return(FALSE)
        if (isTRUE(a) && isTRUE(b)) TRUE else NA
      } else if (op == "or") {
        a <- ev(node$l)
        if (isTRUE(a)) return(TRUE)
        b <- ev(node$r)
        if (isTRUE(b)) return(TRUE)
        if (isFALSE(a) && isFALSE(b)) FALSE else NA
      } else if (op %in% c("=", "!=", "<", "<=", ">", ">=")) {
        side <- NULL
        lit <- NULL
        flip <- FALSE
        if (.sk_is_s(node$l, s) && !is.null(.sk_lit(node$r))) {
          side <- node$l
          lit <- .sk_lit(node$r)
        } else if (.sk_is_s(node$r, s) && !is.null(.sk_lit(node$l))) {
          side <- node$r
          lit <- .sk_lit(node$l)
          flip <- TRUE
        }
        if (is.null(side)) {
          return(NA)
        }
        if (op %in% c("=", "!=")) {
          eq <- identical(v, lit)
          nv <- suppressWarnings(as.numeric(c(v, lit)))
          if (!eq && !anyNA(nv)) eq <- nv[1L] == nv[2L]
          return(if (op == "=") eq else !eq)
        }
        nv <- suppressWarnings(as.numeric(c(v, lit)))
        if (anyNA(nv)) {
          return(NA)
        }
        a <- nv[1L]
        b <- nv[2L]
        if (flip) {
          tmp <- a
          a <- b
          b <- tmp
        }
        switch(op, "<" = a < b, "<=" = a <= b, ">" = a > b, ">=" = a >= b)
      } else {
        NA
      }
    },
    call = {
      if (node$fn == "not" && length(node$args) == 1L) {
        r <- ev(node$args[[1L]])
        if (is.na(r)) NA else !r
      } else if (node$fn == "selected" && length(node$args) == 2L &&
                 .sk_is_s(node$args[[1L]], s) &&
                 !is.null(.sk_lit(node$args[[2L]]))) {
        identical(v, .sk_lit(node$args[[2L]]))
      } else {
        NA
      }
    },
    NA
  )
}

#' Matching end row of every begin row (NA for other rows)
#' @noRd
.sk_ends <- function(kind) {
  n <- length(kind)
  ends <- rep(NA_integer_, n)
  stack <- integer(0)
  for (i in seq_len(n)) {
    if (kind[i] %in% c("begin_group", "begin_repeat")) {
      stack <- c(stack, i)
    } else if (kind[i] %in% c("end_group", "end_repeat") && length(stack)) {
      ends[stack[length(stack)]] <- i
      stack <- stack[-length(stack)]
    }
  }
  ends[is.na(ends) & kind %in% c("begin_group", "begin_repeat")] <- n + 1L
  ends
}

#' Skip target of one choice
#'
#' @return The survey row index of the next asked row, `0L` for the end of
#'   the form, or NA when no instruction applies (nothing is skipped, the
#'   result is unknown, or the skip leaves a repeat).
#' @noRd
.sk_target <- function(sv, trees, refs, ends, s, v) {
  n <- nrow(sv)
  name <- sv$name[s]
  kind <- sv$kind
  skipped <- FALSE
  par <- sv$parent[s]
  cend <- if (is.na(par)) n + 1L else ends[par]
  j <- s + 1L
  repeat {
    if (j >= cend) {
      if (is.na(par)) {
        return(if (skipped) 0L else NA_integer_)
      }
      if (kind[par] == "begin_repeat") {
        return(NA_integer_)
      }
      j <- cend + 1L
      par <- sv$parent[par]
      cend <- if (is.na(par)) n + 1L else ends[par]
      next
    }
    k <- kind[j]
    if (k %in% c("end_group", "end_repeat")) {
      j <- j + 1L
      next
    }
    rel <- TRUE
    if (name %in% refs[[j]]) {
      rel <- if (is.null(trees[[j]])) NA else .sk_eval(trees[[j]], name, v)
    }
    if (is.na(rel)) {
      return(NA_integer_)
    }
    is_begin <- k %in% c("begin_group", "begin_repeat")
    if (!rel) {
      if (k != "hidden") skipped <- TRUE
      j <- if (is_begin) ends[j] + 1L else j + 1L
      next
    }
    if (!is.na(sv$serial[j]) && !(k %in% c("hidden", "note"))) {
      return(if (skipped) j else NA_integer_)
    }
    j <- j + 1L
  }
}

#' Skip instructions for every single-choice question
#'
#' @param form A `cxf_form` whose survey carries the serials in use.
#' @return A list with one element per survey row: NULL, or a character
#'   vector named by choice value giving the target serial (`"0"` for the end
#'   of the form, NA for no instruction).
#' @noRd
.skip_targets <- function(form) {
  sv <- form$survey
  n <- nrow(sv)
  out <- vector("list", n)
  if (n == 0L) {
    return(out)
  }
  rel <- sv$relevant
  has <- !is.na(rel) & grepl("${", rel, fixed = TRUE)
  trees <- vector("list", n)
  refs <- rep(list(character(0)), n)
  for (j in which(has)) {
    tr <- tryCatch(.hx_parse(.hx_tokenize(rel[j])), error = function(e) NULL)
    trees[j] <- list(tr)
    refs[[j]] <- if (is.null(tr)) {
      m <- regmatches(rel[j], gregexpr("\\$\\{[^}]*\\}", rel[j]))[[1L]]
      trimws(sub("^\\$\\{(.*)\\}$", "\\1", m))
    } else {
      .sk_refs(tr)
    }
  }
  all_refs <- unique(unlist(refs, use.names = FALSE))
  cand <- which(sv$kind == "select_one" & !is.na(sv$name) &
                  sv$name %in% all_refs)
  if (length(cand) == 0L) {
    return(out)
  }
  ends <- .sk_ends(sv$kind)
  for (s in cand) {
    ch <- .select_choices(form, s)$choices
    vals <- ch$name
    if (isTRUE(sv$or_other[s])) vals <- c(vals, "other")
    if (length(vals) == 0L) next
    tg <- vapply(vals, function(v) .sk_target(sv, trees, refs, ends, s, v),
                 1L)
    if (all(is.na(tg))) next
    txt <- ifelse(is.na(tg), NA_character_,
                  ifelse(tg == 0L, "0", sv$serial[pmax(tg, 1L)]))
    out[[s]] <- stats::setNames(txt, vals)
  }
  out
}

#' Display text of a skip target
#' @noRd
.skip_text <- function(target, strings) {
  if (is.na(target)) {
    return(NA_character_)
  }
  if (identical(target, "0")) {
    return(strings$go_to_end)
  }
  paste0(strings$go_to, " Q", target)
}
