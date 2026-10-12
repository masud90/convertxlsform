#' Command Line Interface
#'
#' Run via: Rscript -e "convertxlsform::cli()" -- [options]
#' Or, after installation: convertxlsform -- [options]
#'
#' @export
cli <- function() {
  suppressPackageStartupMessages(library(optparse))

  option_list <- list(
    make_option(c("-x", "--xlsform"), type = "character", help = "Path to XLSForm .xlsx [required]"),
    make_option(c("-o", "--out"), type = "character", default = NULL, help = "Output Word file (.docx)"),
    make_option(c("-l", "--lang"), type = "character", default = "en", help = "Language (default: en)"),
    make_option("--no-relevants", action = "store_false", dest = "include_relevants", default = TRUE,
                help = "Do not include skip logics"),
    make_option("--no-choice-name", action = "store_false", dest = "include_choice_name", default = TRUE,
                help = "Do not include choice names"),
    make_option("--no-serials", action = "store_false", dest = "nested_serial_numbers", default = TRUE,
                help = "Do not use nested serial numbers"),
    make_option("--media-root", type = "character", default = ".", help = "Folder with media files"),
    make_option("--quiet", action = "store_false", dest = "verbose", default = TRUE,
                help = "Suppress progress messages")
  )

  parser <- OptionParser(option_list = option_list)
  args <- parse_args(parser, positional_arguments = FALSE)

  if (is.null(args$xlsform)) {
    print_help(parser)
    quit(status = 1)
  }

  convertxlsform(
    xlsform_path = args$xlsform,
    output_docx = args$out,
    lang = args$lang,
    include_relevants = args$include_relevants,
    include_choice_name = args$include_choice_name,
    nested_serial_numbers = args$nested_serial_numbers,
    media_root = args$`media-root`,
    verbose = args$verbose
  )
}
