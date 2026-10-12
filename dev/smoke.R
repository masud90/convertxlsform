# Smoke test: load the package from source and convert the bundled multilang example.
# Run from the package root: Rscript dev/smoke.R

if (requireNamespace("pkgload", quietly = TRUE)) {
  pkgload::load_all(".")
} else {
  install.packages(".", repos = NULL, type = "source")
  library(convertxlsform)
}

example <- system.file("extdata/multilang/multilang.xlsx", package = "convertxlsform")
stopifnot(nzchar(example))

res <- convertxlsform(example, verbose = TRUE)
print(res)
