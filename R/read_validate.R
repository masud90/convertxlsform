# Phase 2: Reading, validation, normalization, language resolution, parsing

# Public-ish (internal) main entry for Phase 2 logic

.ensure_columns <- function(df, cols, type = "character") {
  if (is.null(df)) return(df)
  for (nm in cols) {
    if (!nm %in% names(df)) {
      if (type == "character") df[[nm]] <- NA_character_
      else if (type == "logical") df[[nm]] <- NA
      else if (type == "numeric") df[[nm]] <- NA_real_
      else df[[nm]] <- NA
    }
  }
  df
}

.read_and_validate_xlsform <- function(xlsform_path, lang = "en", verbose = TRUE) {
  .stop_if_missing_file(xlsform_path)

  if (verbose) message("Reading XLSForm sheets…")
  sheets <- .safe_read_sheets(xlsform_path)

  if (verbose) message("Validating sheet presence and required columns…")
  .validate_required_sheets(sheets)
  .validate_required_columns(sheets)

  if (verbose) message("Normalizing column names…")
  sheets$survey  <- .normalize_cols(sheets$survey)
  sheets$choices <- .normalize_cols(sheets$choices)
  if (!is.null(sheets$settings)) sheets$settings <- .normalize_cols(sheets$settings)

  # ensure optional columns exist so downstream mutate() never fails
  # survey optional fields we may touch later
  survey_opt_char <- c(
    "hint", "required_message", "constraint", "constraint_message",
    "appearance", "parameters", "relevant", "image", "audio", "video",
    "repeat_count", "choice_filter"
  )
  sheets$survey  <- .ensure_columns(sheets$survey, survey_opt_char, type = "character")
  # 'required' may be missing; create it so we can coerce to logical in mutate
  if (!"required" %in% names(sheets$survey)) {
    sheets$survey$required <- NA
  }

  # choices optional media fields
  choices_opt_char <- c("image", "audio", "video")
  sheets$choices <- .ensure_columns(sheets$choices, choices_opt_char, type = "character")
  # <<< END NEW

  if (verbose) message("Resolving language columns (label/hint/*_message)…")
  lang_res <- .resolve_language_columns(
    survey  = sheets$survey,
    choices = sheets$choices,
    requested_lang = lang,
    verbose = verbose
  )

  survey  <- lang_res$survey
  choices <- lang_res$choices
  lang_used <- lang_res$lang_used
  lang_mode <- lang_res$lang_mode

  if (verbose) {
    if (lang_mode == "fallback_en") {
      message("  Falling back to English (en).")
    } else if (lang_mode == "single_label") {
      message("  Detected single-language form; using 'label'.")
    } else {
      message("  Using language: ", lang_used)
    }
  }

  if (verbose) message("Parsing question types and parameters…")
  survey <- survey %>%
    dplyr::mutate(
      base_type = .parse_base_type(type),
      type_arg  = .parse_type_arg(type),
      required  = .coerce_logical_maybe(required)
    )

  survey <- survey %>%
    dplyr::mutate(
      parameters_parsed = purrr::map(parameters %||% rep(NA_character_, n()),
                                     .parse_parameters_vectorized)
    )

  list(
    survey   = survey,
    choices  = choices,
    settings = sheets$settings,
    lang_used = lang_used,
    lang_mode = lang_mode
  )
}

# ---------- IO helpers ----------

.expand_path <- function(path) {
  # Expand ~ and relative paths; return absolute path if it exists or the original otherwise
  tryCatch(normalizePath(path, winslash = "/", mustWork = FALSE),
           error = function(e) path)
}

.stop_if_missing_file <- function(path) {
  if (is.null(path)) {
    rlang::abort("Input file path is NULL.")
  }
  apath <- .expand_path(path)
  if (!file.exists(apath)) {
    rlang::abort(paste0(
      "Input file not found: ", path,
      "\n  Resolved/checked at: ", apath,
      "\n  Current working directory: ", getwd()
    ))
  }
}

.safe_read_sheets <- function(path) {
  available <- readxl::excel_sheets(path)
  get_or_null <- function(name) {
    if (name %in% available) readxl::read_excel(path, sheet = name) else NULL
  }
  list(
    survey  = get_or_null("survey"),
    choices = get_or_null("choices"),
    settings= get_or_null("settings")
  )
}

# ---------- Validation ----------

.validate_required_sheets <- function(sheets) {
  missing <- character()
  if (is.null(sheets$survey))  missing <- c(missing, "survey")
  if (is.null(sheets$choices)) missing <- c(missing, "choices")
  # settings optional
  if (length(missing)) {
    rlang::abort(paste0("Missing required sheet(s): ", paste(missing, collapse = ", ")))
  }
}

.validate_required_columns <- function(sheets) {
  # For survey: require type, name, and at least one label column (label or label::<lang>)
  survey_cols <- tolower(names(sheets$survey))
  choices_cols <- tolower(names(sheets$choices))
  # Required survey bare columns
  survey_req <- c("type", "name")
  survey_missing <- setdiff(survey_req, survey_cols)
  if (length(survey_missing)) {
    rlang::abort(paste0("Missing columns in sheet 'survey': ", paste(survey_missing, collapse = ", ")))
  }
  # Survey requires some label variant
  has_any_label <- any(grepl("^label($|::)", tolower(names(sheets$survey))))
  if (!has_any_label) {
    rlang::abort("Missing label column in sheet 'survey': need 'label' or at least one 'label::<lang>'.")
  }

  # For choices: require list_name, name, and at least one label variant
  choices_req <- c("list_name", "name")
  choices_missing <- setdiff(choices_req, choices_cols)
  if (length(choices_missing)) {
    rlang::abort(paste0("Missing columns in sheet 'choices': ", paste(choices_missing, collapse = ", ")))
  }
  has_any_choices_label <- any(grepl("^label($|::)", tolower(names(sheets$choices))))
  if (!has_any_choices_label) {
    rlang::abort("Missing label column in sheet 'choices': need 'label' or at least one 'label::<lang>'.")
  }

  # Settings: optional. If present and 'form_title' exists, later phases will use it.
  if (!is.null(sheets$settings)) {
    # No hard requirement here yet; we’ll enforce title usage in composition phase.
    # But we can sanity check if present: 1 non-NA value
    if ("form_title" %in% tolower(names(sheets$settings))) {
      # ok
    }
  }
}

# ---------- Normalization ----------

.normalize_cols <- function(df) {
  if (is.null(df)) return(NULL)
  names(df) <- tolower(names(df))
  df
}

# ---------- Language resolution ----------

.resolve_language_columns <- function(survey, choices, requested_lang = "en", verbose = TRUE) {
  # Normalize requested_lang input; accept "English (en)" and "en"
  req_lang <- .normalize_lang(requested_lang)

  # Given columns: choose label/hint/*_message columns for survey and choices
  # Selection order:
  # 1) label::<req_lang> if exists
  # 2) plain label (single-language mode)
  # 3) label::en (fallback to English (en))
  # Otherwise -> error.
  sel_survey <- .select_language_set(survey, req_lang, context = "survey")
  sel_choices <- .select_language_set(choices, req_lang, context = "choices")

  lang_mode <- sel_survey$lang_mode
  lang_used <- sel_survey$lang_used
  # For consistency, if survey fell back and choices didn't, we still report the survey decision.

  survey2 <- survey %>%
    dplyr::mutate(
      label_selected = sel_survey$label,
      hint_selected  = sel_survey$hint,
      reqmsg_selected= sel_survey$reqmsg,
      cmsg_selected  = sel_survey$cmsg
    )

  choices2 <- choices %>%
    dplyr::mutate(
      label_selected = sel_choices$label
    )

  list(
    survey   = survey2,
    choices  = choices2,
    lang_used= lang_used,
    lang_mode= lang_mode
  )
}

.normalize_lang <- function(lang) {
  lang <- trimws(tolower(lang %||% "en"))
  if (lang == "english (en)") return("en")
  if (lang == "english") return("en")
  lang
}

.select_language_set <- function(df, req_lang, context = "survey") {
  cols <- names(df)

  pick_col <- function(prefix) {
    # Returns a vector of the selected column, or NA_character_
    # Priority: ::req_lang, then plain, then ::en
    cand_req <- paste0(prefix, "::", req_lang)
    cand_plain <- prefix
    cand_en <- paste0(prefix, "::en")

    if (cand_req %in% cols) {
      mode <- "multilingual"; used <- req_lang
      return(list(vec = df[[cand_req]], mode = mode, used = used))
    }
    if (cand_plain %in% cols) {
      mode <- "single_label"; used <- "plain"
      return(list(vec = df[[cand_plain]], mode = mode, used = used))
    }
    if (cand_en %in% cols) {
      mode <- "fallback_en"; used <- "en"
      return(list(vec = df[[cand_en]], mode = mode, used = used))
    }
    list(vec = rep(NA_character_, nrow(df)), mode = "missing", used = NA_character_)
  }

  lab <- pick_col("label")
  hnt <- pick_col("hint")
  rqm <- pick_col("required_message")
  csm <- pick_col("constraint_message")

  # At minimum, label must not be missing entirely
  if (all(is.na(lab$vec))) {
    rlang::abort(
      paste0("Could not find any usable label column in sheet '", context,
             "'. Tried label::", req_lang, ", label, and label::en.")
    )
  }

  lang_mode <- lab$mode
  lang_used <- lab$used

  list(
    label    = lab$vec,
    hint     = hnt$vec,
    reqmsg   = rqm$vec,
    cmsg     = csm$vec,
    lang_mode= lang_mode,
    lang_used= lang_used
  )
}

# ---------- Type & parameters parsing ----------

.parse_base_type <- function(type_vec) {
  type_vec <- as.character(type_vec %||% NA_character_)
  # Take the part before the first space or '['
  bt <- stringr::str_trim(stringr::str_extract(type_vec, "^[^\\s\\[]+"))
  # Normalize common aliases
  bt <- dplyr::case_when(
    tolower(bt) %in% c("datetime", "dateTime") ~ "datetime",
    TRUE ~ tolower(bt)
  )
  bt
}

.parse_type_arg <- function(type_vec) {
  type_vec <- as.character(type_vec %||% NA_character_)
  # extract content inside first [ ... ]
  arg <- stringr::str_match(type_vec, "\\[(.*?)\\]")
  if (nrow(arg)) {
    return(arg[,2])
  }
  NA_character_
}

.parse_parameters_vectorized <- function(x) {
  if (is.na(x) || !nzchar(x)) return(rlang::list2())
  # Expect "start=1 end=10 step=1" style; handle commas or spaces
  tokens <- unlist(strsplit(x, "[,\\s]+"))
  tokens <- tokens[nzchar(tokens)]
  kv <- purrr::map(tokens, function(tk) {
    if (!grepl("=", tk, fixed = TRUE)) return(NULL)
    parts <- strsplit(tk, "=", fixed = TRUE)[[1]]
    if (length(parts) != 2) return(NULL)
    key <- trimws(tolower(parts[1]))
    val <- trimws(parts[2])
    # Try numeric; else keep as string
    val_num <- suppressWarnings(as.numeric(val))
    if (!is.na(val_num)) val <- val_num
    rlang::set_names(list(val), key)
  })
  kv <- kv[!vapply(kv, is.null, logical(1))]
  # collapse into a named list
  out <- list()
  for (i in seq_along(kv)) out[names(kv[[i]])] <- kv[[i]]
  out
}

# ---------- Logical coercion ----------

.coerce_logical_maybe <- function(x) {
  if (is.null(x)) return(rep(NA, 0))
  # Accept yes/true/1 (case-insensitive)
  y <- tolower(as.character(x))
  dplyr::case_when(
    y %in% c("1", "true", "yes", "y", "t") ~ TRUE,
    y %in% c("0", "false", "no", "n", "f") ~ FALSE,
    TRUE ~ NA
  )
}

# ---------- Utilities ----------

`%||%` <- function(a, b) if (!is.null(a)) a else b
