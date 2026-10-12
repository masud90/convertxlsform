# A stand-in for .render_docx() so path logic can be tested without the
# renderer. Works before the renderer exists (assigns into the namespace)
# and after (mocks the binding).
stub_render_docx <- function(env = parent.frame()) {
  stub <- function(model, opts, path) {
    officer::read_docx() %>%
      officer::body_add_par("stub") %>%
      print(target = path)
    invisible(path)
  }
  testthat::local_mocked_bindings(.render_docx = stub, .package = "convertxlsform",
                                  .env = env)
  invisible(NULL)
}

example_form <- function() {
  system.file("extdata", "multilang", "multilang.xlsx",
              package = "convertxlsform")
}

# A writable copy of the example in a fresh folder
copy_example <- function(env = parent.frame()) {
  dir <- withr::local_tempdir(.local_envir = env)
  src_dir <- dirname(example_form())
  file.copy(list.files(src_dir, full.names = TRUE), dir)
  dest <- file.path(dir, "myform.xlsx")
  file.rename(file.path(dir, "multilang.xlsx"), dest)
  dest
}
