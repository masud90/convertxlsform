# Phase 3: Group/Repeat parsing & nested serial numbers

# Public-ish (internal) main entry for Phase 3 logic
.build_structure <- function(survey, nested_serial_numbers = TRUE, verbose = TRUE) {
  if (verbose) message("Building structure (groups/repeats, depth, numbering)…")

  # Safety: ensure columns referenced exist
  survey <- .ensure_columns(survey, c("label_selected", "repeat_count", "type", "name"), type = "character")
  survey <- .ensure_columns(survey, c("base_type", "type_arg"), type = "character")

  # Stack entries will be lists with: kind ("group"|"repeat"), label, row_index, counter (for numbering)
  stack <- list()
  # For nested numbering, keep counters per depth
  counters <- integer(0)

  items <- vector("list", nrow(survey))

  # Helper: push to stack
  push <- function(x) { stack <<- c(stack, list(x)) }
  # Helper: pop from stack (error if empty)
  pop  <- function(expected_kind, end_row) {
    if (length(stack) == 0L) {
      rlang::abort(paste0("Unmatched ", expected_kind, " terminator at survey row ", end_row, ". No corresponding begin_*."))
    }
    top <- stack[[length(stack)]]
    stack <<- stack[-length(stack)]
    top
  }
  # Helper: current depth
  depth <- function() length(stack)

  # Numbering helpers
  bump_counter <- function() {
    d <- depth()
    # Expand counters vector if needed
    if (length(counters) < d + 1L) {
      counters <<- c(counters, rep(0L, d + 1L - length(counters)))
    }
    # Increment counter at current depth
    counters[d + 1L] <<- counters[d + 1L] + 1L
    # Zero-out deeper levels (if any) when we step to a sibling
    if (length(counters) > d + 1L) {
      counters[(d + 2L):length(counters)] <<- 0L
    }
    counters[seq_len(d + 1L)]
  }
  get_serial <- function() {
    paste0(bump_counter(), collapse = ".")
  }
  reset_counter_for_new_scope <- function() {
    # When we enter a new scope (begin_*), ensure counter space exists but don't increment yet.
    d <- depth() + 1L
    if (length(counters) < d) {
      counters <<- c(counters, rep(0L, d - length(counters)))
    }
    # For a new scope, deeper levels should be zero
    if (length(counters) > d) {
      counters[(d + 1L):length(counters)] <<- 0L
    }
  }

  # Define which base_types are control markers or ignored
  is_control_begin <- function(bt) bt %in% c("begin_group", "begin_repeat")
  is_control_end   <- function(bt) bt %in% c("end_group", "end_repeat")
  is_ignored_type  <- function(bt) bt %in% c("calculate", "hidden", "xml-external") # fully ignored later

  # Iterate
  for (i in seq_len(nrow(survey))) {
    row <- survey[i, , drop = FALSE]
    bt  <- tolower(as.character(row$base_type %||% row$type %||% ""))

    if (is_control_begin(bt)) {
      kind <- if (bt == "begin_group") "group" else "repeat"
      push(list(kind = kind,
                label = as.character(row$label_selected %||% ""),
                row_index = i,
                name = as.character(row$name %||% NA_character_),
                repeat_count = as.character(row$repeat_count %||% NA_character_)))
      # entering new scope: prep numbering spacing
      reset_counter_for_new_scope()

      items[[i]] <- list(
        idx = i,
        kind = paste0("begin_", kind),
        base_type = bt,
        type_arg = as.character(row$type_arg %||% NA_character_),
        depth = depth(), # after push
        label = as.character(row$label_selected %||% ""),
        inherited_label = as.character(row$label_selected %||% ""), # for convenience
        repeat_count = if (kind == "repeat") as.character(row$repeat_count %||% NA_character_) else NA_character_,
        is_question = FALSE,
        serial = NA_character_
      )
      next
    }

    if (is_control_end(bt)) {
      expected_kind <- if (bt == "end_group") "group" else "repeat"
      top <- pop(expected_kind, end_row = i)
      if (top$kind != expected_kind) {
        rlang::abort(paste0(
          "Mismatched end marker at survey row ", i, ". Expected end_", top$kind, " to match begin at row ", top$row_index, "."
        ))
      }
      items[[i]] <- list(
        idx = i,
        kind = paste0("end_", expected_kind),
        base_type = bt,
        type_arg = NA_character_,
        depth = depth(), # depth BEFORE pop (we popped already; report current)
        label = NA_character_,            # ignore end row's own label
        inherited_label = as.character(top$label %||% ""),
        repeat_count = if (expected_kind == "repeat") as.character(top$repeat_count %||% NA_character_) else NA_character_,
        is_question = FALSE,
        serial = NA_character_
      )
      next
    }

    # Non-control rows
    if (is_ignored_type(bt)) {
      # Produce a placeholder item we can filter out later if needed
      items[[i]] <- list(
        idx = i,
        kind = "ignored",
        base_type = bt,
        type_arg = as.character(row$type_arg %||% NA_character_),
        depth = depth(),
        label = as.character(row$label_selected %||% ""),
        inherited_label = NA_character_,
        repeat_count = NA_character_,
        is_question = FALSE,
        serial = NA_character_
      )
      next
    }

    # Question-like rows (including "note")
    serial <- if (nested_serial_numbers) get_serial() else NA_character_

    items[[i]] <- list(
      idx = i,
      kind = "question",
      base_type = bt,
      type_arg = as.character(row$type_arg %||% NA_character_),
      depth = depth(),
      label = as.character(row$label_selected %||% ""),
      inherited_label = NA_character_,
      repeat_count = NA_character_,
      is_question = TRUE,
      serial = serial,
      name = as.character(row$name %||% NA_character_)
    )
  }

  # After iteration, stack must be empty
  if (length(stack) > 0L) {
    # Build an informative message about the first unclosed begin
    top <- stack[[length(stack)]]
    rlang::abort(paste0(
      "Unclosed begin_", top$kind, " starting at survey row ", top$row_index, ". Missing a matching end_ marker."
    ))
  }

  # Bind list to tibble
  out <- tibble::as_tibble(purrr::list_rbind(items))

  if (verbose) {
    n_q <- sum(out$is_question, na.rm = TRUE)
    n_bg <- sum(out$kind == "begin_group")
    n_br <- sum(out$kind == "begin_repeat")
    message("  Structure built: ",
            nrow(out), " rows (questions=", n_q, ", begin_group=", n_bg, ", begin_repeat=", n_br, ")")
  }

  out
}
