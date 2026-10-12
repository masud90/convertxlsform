# Humanizing XLSForm expressions (relevant, constraint, choice_filter,
# repeat_count) for the paper form. ASCII-only source.
#
# A small tokenizer and recursive-descent parser turn the expression into a
# tree, which is then rendered as readable text. Anything the parser or the
# renderer does not recognize (unknown functions, paths, odd syntax) makes the
# whole expression fall back to the original text with only `${name}`
# substituted.

#' Tokenize an expression
#'
#' Token types: `ref` (`${name}`), `str`, `num`, `op` (including the words
#' `and`, `or`, `div`, `mod`), `func` (a name followed by an opening
#' parenthesis), `ident` (a bare word), `lparen`, `rparen`, `comma`, `dot`.
#' Aborts on any character it does not know.
#'
#' @param s A single expression string.
#' @return A list of tokens, each `list(type, value)`.
#' @noRd
.hx_tokenize <- function(s) {
  toks <- list()
  rest <- s
  fail <- function() stop("cannot tokenize", call. = FALSE)
  m1 <- function(pat) {
    r <- regexec(pat, rest, perl = TRUE)[[1L]]
    if (r[1L] == -1L) {
      return(NULL)
    }
    regmatches(rest, list(r))[[1L]]
  }
  push <- function(type, value, consumed) {
    toks[[length(toks) + 1L]] <<- list(type = type, value = value)
    rest <<- substring(rest, consumed + 1L)
  }
  op_words <- c("and", "or", "div", "mod")
  repeat {
    rest <- sub("^[[:space:]]+", "", rest)
    if (!nzchar(rest)) {
      break
    }
    if (!is.null(m <- m1("^\\$\\{[[:space:]]*([^}]*?)[[:space:]]*\\}"))) {
      if (!nzchar(m[2L])) fail()
      push("ref", m[2L], nchar(m[1L]))
    } else if (!is.null(m <- m1("^'([^']*)'"))) {
      push("str", m[2L], nchar(m[1L]))
    } else if (!is.null(m <- m1("^\"([^\"]*)\""))) {
      push("str", m[2L], nchar(m[1L]))
    } else if (!is.null(m <- m1("^([0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)"))) {
      push("num", m[2L], nchar(m[1L]))
    } else if (!is.null(m <- m1("^(!=|>=|<=|=|>|<|\\+|-|\\*)"))) {
      push("op", m[2L], nchar(m[1L]))
    } else if (!is.null(m <- m1("^[A-Za-z_][A-Za-z0-9_.-]*"))) {
      w <- m[1L]
      after <- substring(rest, nchar(w) + 1L)
      if (w %in% op_words) {
        push("op", w, nchar(w))
      } else if (grepl("^[[:space:]]*\\(", after)) {
        push("func", w, nchar(w))
      } else {
        push("ident", w, nchar(w))
      }
    } else if (startsWith(rest, "(")) {
      push("lparen", "(", 1L)
    } else if (startsWith(rest, ")")) {
      push("rparen", ")", 1L)
    } else if (startsWith(rest, ",")) {
      push("comma", ",", 1L)
    } else if (startsWith(rest, ".")) {
      push("dot", ".", 1L)
    } else {
      fail()
    }
  }
  toks
}

#' Parse tokens into a tree
#'
#' Precedence, lowest first: `or`, `and`, comparisons, `+ -`, `* div mod`,
#' unary minus. Aborts on a syntax error or leftover tokens.
#' @noRd
.hx_parse <- function(toks) {
  pos <- 1L
  n <- length(toks)
  fail <- function() stop("cannot parse", call. = FALSE)
  peek <- function() if (pos <= n) toks[[pos]] else NULL
  adv <- function() {
    if (pos > n) fail()
    t <- toks[[pos]]
    pos <<- pos + 1L
    t
  }
  is_op <- function(t, ops) !is.null(t) && t$type == "op" && t$value %in% ops
  expect <- function(type) {
    t <- peek()
    if (is.null(t) || t$type != type) fail()
    adv()
  }
  left_assoc <- function(sub, ops) {
    force(sub)
    force(ops)
    function() {
      l <- sub()
      while (is_op(peek(), ops)) {
        o <- adv()$value
        r <- sub()
        l <- list(t = "bin", op = o, l = l, r = r)
      }
      l
    }
  }
  parse_primary <- function() {
    t <- adv()
    switch(
      t$type,
      ref = list(t = "ref", name = t$value),
      str = list(t = "str", value = t$value),
      num = list(t = "num", text = t$value),
      dot = list(t = "dot"),
      ident = list(t = "ident", text = t$value),
      func = {
        expect("lparen")
        args <- list()
        if (!is.null(peek()) && peek()$type == "rparen") {
          adv()
        } else {
          repeat {
            args[[length(args) + 1L]] <- parse_or()
            nt <- peek()
            if (!is.null(nt) && nt$type == "comma") {
              adv()
            } else {
              break
            }
          }
          expect("rparen")
        }
        list(t = "call", fn = t$value, args = args)
      },
      lparen = {
        x <- parse_or()
        expect("rparen")
        list(t = "paren", x = x)
      },
      fail()
    )
  }
  parse_unary <- function() {
    if (is_op(peek(), "-")) {
      adv()
      return(list(t = "un", op = "-", x = parse_unary()))
    }
    parse_primary()
  }
  parse_mul <- left_assoc(function() parse_unary(), c("*", "div", "mod"))
  parse_add <- left_assoc(function() parse_mul(), c("+", "-"))
  parse_cmp <- left_assoc(function() parse_add(),
                          c("=", "!=", "<", ">", "<=", ">="))
  parse_and <- left_assoc(function() parse_cmp(), "and")
  parse_or <- left_assoc(function() parse_and(), "or")

  if (n == 0L) fail()
  tree <- parse_or()
  if (pos <= n) fail()
  tree
}

#' Lookup context for humanizing: survey names, serials, choice labels
#' @noRd
.expr_ctx <- function(form, strings = NULL, show_variable_names = TRUE) {
  list(form = form, sv = form$survey, cache = new.env(parent = emptyenv()),
       s = strings %||% .cxf_strings(.form_lang_code(form)),
       show_names = isTRUE(show_variable_names))
}

#' Survey row index of a question name (NA if absent)
#' @noRd
.ctx_row <- function(ctx, name) {
  match(name, ctx$sv$name)
}

#' Choices (data.frame with name, label) of the question at row `i`, or NULL
#' @noRd
.ctx_choices <- function(ctx, i) {
  if (is.na(i)) {
    return(NULL)
  }
  key <- as.character(i)
  if (!is.null(hit <- ctx$cache[[key]])) {
    return(hit$value)
  }
  sv <- ctx$sv
  arg <- sv$type_arg[i]
  ch <- NULL
  if (sv$kind[i] %in% c("select_one", "select_multiple", "rank") &&
      !is.na(arg)) {
    bt <- sv$base_type[i]
    ch <- if (grepl("_from_file$", bt)) {
      .file_choices(ctx$form, arg)$choices
    } else if (grepl("_external$", bt)) {
      NULL
    } else {
      .choice_list(ctx$form, arg)
    }
  }
  assign(key, list(value = ch), envir = ctx$cache)
  ch
}

#' Label of a choice value for the question at row `i` (value if unknown)
#' @noRd
.ctx_label <- function(ctx, i, value) {
  ch <- .ctx_choices(ctx, i)
  if (is.null(ch) || nrow(ch) == 0L) {
    return(value)
  }
  j <- match(value, ch$name)
  if (is.na(j) || .is_blank(ch$label[j])) value else ch$label[j]
}

#' Display text of a `${name}` reference
#' @noRd
.hx_ref <- function(name, ctx) {
  i <- .ctx_row(ctx, name)
  if (is.na(i) || is.na(ctx$sv$serial[i])) {
    return(name)
  }
  if (ctx$show_names) {
    paste0("Q", ctx$sv$serial[i], " (", name, ")")
  } else {
    paste0("Q", ctx$sv$serial[i])
  }
}

#' Quote a string literal for display
#' @noRd
.hx_quote <- function(v, empty = "empty") {
  if (!nzchar(v)) {
    return(empty)
  }
  if (grepl("\"", v, fixed = TRUE)) paste0("'", v, "'") else paste0("\"", v, "\"")
}

#' Original text with only `${name}` substituted
#' @noRd
.hx_subst <- function(expr, ctx) {
  m <- gregexpr("\\$\\{[^}]*\\}", expr)
  refs <- regmatches(expr, m)[[1L]]
  if (length(refs) == 0L) {
    return(trimws(expr))
  }
  nm <- trimws(sub("^\\$\\{(.*)\\}$", "\\1", refs))
  new <- vapply(nm, function(x) .hx_ref(x, ctx), "", USE.NAMES = FALSE)
  regmatches(expr, m) <- list(new)
  trimws(expr)
}

#' Survey row of the question a node refers to (NA for anything else)
#' @noRd
.hx_node_row <- function(node, ctx, cur) {
  if (node$t == "ref") {
    .ctx_row(ctx, node$name)
  } else if (node$t == "dot") {
    cur
  } else {
    NA_integer_
  }
}

#' Render a tree as text
#' @noRd
.hx_render <- function(n, ctx, field, cur) {
  r <- function(x) .hx_render(x, ctx, field, cur)
  switch(
    n$t,
    ref = .hx_ref(n$name, ctx),
    str = .hx_quote(n$value, ctx$s$empty),
    num = n$text,
    dot = if (field == "constraint") ctx$s$the_answer else ctx$s$this_answer,
    ident = n$text,
    paren = paste0("(", r(n$x), ")"),
    un = paste0("-", r(n$x)),
    call = .hx_call(n, ctx, field, cur),
    bin = .hx_bin(n, ctx, field, cur),
    stop("unknown node", call. = FALSE)
  )
}

#' Render `a op b`
#' @noRd
.hx_bin <- function(n, ctx, field, cur, negate = FALSE) {
  r <- function(x) .hx_render(x, ctx, field, cur)
  op <- n$op
  if (negate) {
    op <- c("=" = "!=", "!=" = "=")[[op]]
  }
  is_lit <- function(x) x$t %in% c("str", "num")
  if (op %in% c("=", "!=") && (is_lit(n$r) || is_lit(n$l))) {
    # Put the literal on the right.
    side <- if (is_lit(n$r)) n$l else n$r
    lit <- if (is_lit(n$r)) n$r else n$l
    if (lit$t == "str" && !nzchar(lit$value)) {
      return(paste(r(side),
                   if (op == "=") ctx$s$is_empty else ctx$s$is_not_empty))
    }
    i <- .hx_node_row(side, ctx, cur)
    if (!is.na(i) && ctx$sv$kind[i] == "select_one") {
      val <- if (lit$t == "str") lit$value else lit$text
      return(paste(r(side), if (op == "=") ctx$s$is else ctx$s$is_not,
                   .hx_quote(.ctx_label(ctx, i, val), ctx$s$empty)))
    }
  }
  txt <- switch(op, "!=" = ctx$s$is_not, div = "/", and = ctx$s$and,
                or = ctx$s$or, op)
  paste(r(n$l), txt, r(n$r))
}

#' Render `selected(a, b)`, optionally negated
#' @noRd
.hx_selected <- function(n, ctx, field, cur, negate = FALSE) {
  if (length(n$args) != 2L) stop("bad arity", call. = FALSE)
  a <- n$args[[1L]]
  b <- n$args[[2L]]
  verb <- if (negate) ctx$s$does_not_include else ctx$s$includes
  rhs <- if (b$t == "str") {
    i <- .hx_node_row(a, ctx, cur)
    .hx_quote(.ctx_label(ctx, i, b$value), ctx$s$empty)
  } else {
    .hx_render(b, ctx, field, cur)
  }
  paste(.hx_render(a, ctx, field, cur), verb, rhs)
}

#' Render a function call
#' @noRd
.hx_call <- function(n, ctx, field, cur) {
  fn <- n$fn
  args <- n$args
  na <- length(args)
  switch(
    fn,
    selected = .hx_selected(n, ctx, field, cur),
    "count-selected" = {
      if (na != 1L) stop("bad arity", call. = FALSE)
      paste(ctx$s$number_selected_in, .hx_render(args[[1L]], ctx, field, cur))
    },
    not = {
      if (na != 1L) stop("bad arity", call. = FALSE)
      a <- args[[1L]]
      if (a$t == "call" && a$fn == "selected") {
        .hx_selected(a, ctx, field, cur, negate = TRUE)
      } else if (a$t == "bin" && a$op %in% c("=", "!=")) {
        .hx_bin(a, ctx, field, cur, negate = TRUE)
      } else {
        paste0(ctx$s$not, " (", .hx_render(a, ctx, field, cur), ")")
      }
    },
    today = , now = , "true" = , "false" = {
      if (na != 0L) stop("bad arity", call. = FALSE)
      switch(fn, today = ctx$s$today, now = ctx$s$now, "true" = ctx$s$true,
             "false" = ctx$s$false, fn)
    },
    stop("unknown function", call. = FALSE)
  )
}

#' Humanize one expression
#' @noRd
.hx_one <- function(expr, ctx, field, cur) {
  if (is.na(expr) || !nzchar(trimws(expr))) {
    return(list(text = NA_character_, fallback = FALSE))
  }
  res <- tryCatch(
    .hx_render(.hx_parse(.hx_tokenize(expr)), ctx, field, cur),
    error = function(e) NULL
  )
  if (is.null(res)) {
    list(text = .hx_subst(expr, ctx), fallback = TRUE)
  } else {
    list(text = res, fallback = FALSE)
  }
}

#' Check the `field` argument
#' @noRd
.check_expr_field <- function(field) {
  ok <- c("relevant", "constraint", "choice_filter", "repeat_count")
  if (!(is.character(field) && length(field) == 1L && field %in% ok)) {
    .cxf_abort("{.arg field} must be one of {.val {ok}}.")
  }
  field
}

#' Humanize an XLSForm expression
#'
#' @param expr A single expression (NA gives NA).
#' @param form A `cxf_form` (its serials and choices are used for lookups).
#' @param field One of `"relevant"`, `"constraint"`, `"choice_filter"`,
#'   `"repeat_count"`. Changes how a bare `.` is shown.
#' @param row Optional survey row index of the question the expression belongs
#'   to, used to find the choice labels for `.`.
#' @param strings Optional string table (see `.cxf_strings()`); by default the
#'   one for the form's language.
#' @param show_variable_names Show `Q3.2 (name)` (TRUE) or `Q3.2` (FALSE) for
#'   references to numbered questions.
#' @return A character string.
#' @noRd
.humanize_expr <- function(expr, form, field = "relevant",
                           row = NA_integer_, strings = NULL,
                           show_variable_names = TRUE) {
  field <- .check_expr_field(field)
  ctx <- .expr_ctx(form, strings, show_variable_names)
  .hx_one(as.character(expr)[1L], ctx, field, row)$text
}

#' Vectorized `.humanize_expr()`
#'
#' @param exprs Character vector.
#' @param rows Optional integer vector (same length) of survey row indices.
#' @param with_fallback If TRUE, return `list(text, fallback)` where `fallback`
#'   flags the expressions that could not be humanized and are shown as
#'   written (with `${name}` substituted).
#' @inheritParams .humanize_expr
#' @return Character vector the same length as `exprs`.
#' @noRd
.humanize_exprs <- function(exprs, form, field = "relevant", rows = NULL,
                            strings = NULL, show_variable_names = TRUE,
                            with_fallback = FALSE) {
  field <- .check_expr_field(field)
  exprs <- as.character(exprs)
  rows <- rows %||% rep(NA_integer_, length(exprs))
  ctx <- .expr_ctx(form, strings, show_variable_names)
  res <- lapply(seq_along(exprs), function(k) {
    .hx_one(exprs[k], ctx, field, rows[k])
  })
  text <- vapply(res, function(x) x$text, "")
  if (!with_fallback) {
    return(text)
  }
  list(text = text, fallback = vapply(res, function(x) x$fallback, NA))
}

#' Plain-language range from a simple numeric constraint
#'
#' Understands `. >= a`, `. < b`, and two such comparisons joined by `and`
#' (the dot on the left, a number on the right). Anything else gives NA.
#'
#' @param constraint A single constraint expression.
#' @param strings A string table from `.cxf_strings()` (English by default).
#' @return A string such as `"from 0 to 30"`, or NA.
#' @noRd
.simple_range_text <- function(constraint, strings = .cxf_strings("en")) {
  if (is.na(constraint) || !nzchar(trimws(constraint))) {
    return(NA_character_)
  }
  tree <- tryCatch(.hx_parse(.hx_tokenize(constraint)),
                   error = function(e) NULL)
  if (is.null(tree)) {
    return(NA_character_)
  }
  cmp_ok <- function(x) {
    x$t == "bin" && x$op %in% c(">=", ">", "<=", "<") && x$l$t == "dot" &&
      x$r$t == "num"
  }
  parts <- if (tree$t == "bin" && tree$op == "and") {
    list(tree$l, tree$r)
  } else {
    list(tree)
  }
  if (!all(vapply(parts, cmp_ok, NA))) {
    return(NA_character_)
  }
  ops <- vapply(parts, function(x) x$op, "")
  val <- vapply(parts, function(x) x$r$text, "")
  if (anyDuplicated(substr(ops, 1L, 1L)) > 0L) {
    return(NA_character_)
  }
  phrase <- c(">=" = strings$at_least, ">" = strings$more_than,
              "<=" = strings$at_most, "<" = strings$less_than)
  if (length(ops) == 2L && setequal(ops, c(">=", "<="))) {
    return(paste(strings$range_from, val[ops == ">="], strings$range_to,
                 val[ops == "<="]))
  }
  paste(paste(phrase[ops], val), collapse = paste0(" ", strings$and, " "))
}
