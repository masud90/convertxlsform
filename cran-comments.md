## Test environments

* Local: macOS 27.0.1 (aarch64-apple-darwin23), R 4.6.0, `R CMD check --as-cran`
  run with `NOT_CRAN=false`, `_R_CHECK_FORCE_SUGGESTS_=false`, `--ignore-vignettes`.
  The package was built with `--no-build-vignettes` because pandoc is not
  installed on this machine.
* GitHub Actions (`R-CMD-check`: macOS, Windows, Ubuntu): not yet run on the
  release candidate. Results to be added before submission.

## R CMD check results

Local run (with and without `--no-manual`): 0 errors | 0 warnings | 3 notes

* CRAN incoming feasibility: "New submission". The development version
  number 0.2.1.9000 also triggers "Version contains large components"; the
  submitted version will be 0.3.0.
* "Package has a VignetteBuilder field but no prebuilt vignette index" and
  "Files 'README.md' or 'NEWS.md' cannot be checked without 'pandoc' being
  installed": both come from the missing pandoc on the local machine.
* "Skipping checking HTML validation: 'tidy' doesn't look like recent enough
  HTML Tidy": the local tidy is the 2006 macOS build (only with the manual).

## Notes for the reviewer

* This is a new release. The package has not been on CRAN before.
* LibreOffice is an optional system requirement, used only to produce PDF. All
  examples and tests that need it are guarded by `has_libreoffice()` and tests
  skip on CRAN. Microsoft Word through the suggested package 'doconv' is a
  fallback.
* Tests take about 38 s with `NOT_CRAN=false` locally.
* Examples and tests write only to `tempdir()`.
