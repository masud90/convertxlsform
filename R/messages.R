# Internal message helpers. User-facing text goes through cli.
# Each wrapper forwards the caller's environment so that {glue} references
# in the messages resolve in the function that raised them.

#' Inform the user, only when verbose
#'
#' @param verbose Logical. Messages are shown only when TRUE.
#' @param ... Message lines (character, named for bullets), passed to
#'   [cli::cli_inform()].
#' @return Invisible NULL.
#' @noRd
.cxf_inform <- function(verbose, ...) {
  if (isTRUE(verbose)) {
    cli::cli_inform(c(...), .envir = parent.frame())
  }
  invisible(NULL)
}

#' Warn the user
#'
#' @param ... Message lines (character, named for bullets), passed to
#'   [cli::cli_warn()].
#' @return Invisible NULL.
#' @noRd
.cxf_warn <- function(...) {
  cli::cli_warn(c(...), .envir = parent.frame())
  invisible(NULL)
}

#' Abort with a cli-formatted error
#'
#' @param ... Message lines (character, named for bullets), passed to
#'   [cli::cli_abort()].
#' @param call The call reported in the error. Defaults to the function that
#'   called this helper.
#' @return Does not return.
#' @noRd
.cxf_abort <- function(..., call = rlang::caller_env()) {
  cli::cli_abort(c(...), call = call, .envir = parent.frame())
}
