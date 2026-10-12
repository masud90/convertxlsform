## R CMD check results

Local run on 2026-10-12, built and checked from a temporary directory with
vignettes and the PDF manual:

```
R CMD build convertxlsform
R CMD check --as-cran convertxlsform_0.3.0.tar.gz
```

0 errors | 0 warnings | 2 notes

* checking CRAN incoming feasibility ... NOTE

  ```
  Maintainer: 'Masud Rahman <write2masud+convertxlsform@gmail.com>'

  New submission
  ```

  This is a new submission.

* checking HTML version of manual ... NOTE

  ```
  Skipping checking HTML validation: 'tidy' doesn't look like recent enough HTML Tidy.
  ```

  The HTML Tidy on the local macOS machine is the 2006 system build. It is not
  a problem with the package.

Results from other platforms: [to fill after win-builder, mac builder, R-hub and
GitHub Actions]

## Test environments

Run locally:

* macOS 27.0.1 (aarch64-apple-darwin23), R 4.6.0, `R CMD check --as-cran`,
  with vignettes (pandoc 3.10 from RStudio) and the PDF manual.

Other platforms:

* GitHub Actions `R-CMD-check` on macOS (release), Windows (release) and Ubuntu
  (devel, release, oldrel-1), commit f81b5b6 on `main`, 2026-10-12: all five
  jobs passed.
* win-builder, R-devel (Windows Server 2022, x86_64): 2 NOTEs: new submission
  and a spelling flag for 'docx' (fixed by quoting); example timing over 10 s
  (fixed by a smaller example form). To be re-run.
* win-builder, R-release (R 4.6.1, Windows Server 2022, x86_64), 2026-10-12,
  after the fixes above: 1 NOTE (new submission). Examples OK; tests 81 s.
* R-hub v2 (`rhub::rhub_check()`), platforms: [to fill after R-hub]
* mac builder (<https://mac.r-project.org/macbuilder/submit.html>): [to fill
  after mac builder]

## Notes for the reviewer

* This is a new submission. The package has not been on CRAN before.
* **Output location.** `convertxlsform(xlsform, output = NULL)` writes the Word
  or PDF file next to the input file the user passes, with the same base name
  and a new extension. The location is derived from a path the user supplied,
  it is reported in a message (and returned invisibly), and the `output`
  argument lets the user choose a file or a folder. An existing file is never
  overwritten unless `overwrite = TRUE`. The package never writes to the home
  directory or the working directory on its own. All examples, tests and the
  vignette write only to `tempdir()`.
* **SystemRequirements.** 'LibreOffice' is optional and is used only to
  produce PDF output. Word output needs only R. The examples that need
  'LibreOffice' run only in interactive sessions and are guarded by
  `has_libreoffice()`. Tests that need it skip when it is not found and on
  CRAN. Microsoft Word through the suggested package 'doconv' is a fallback.
* The tests take about 12 s locally with `NOT_CRAN=false`. The heavy
  end-to-end rendering tests are skipped on CRAN with `skip_on_cran()`; the
  full set takes about 50 s.
* `cli()` prints its help and version text with `writeLines()` because that
  output is the purpose of the terminal interface. Progress messages use
  `message()` and are suppressed with `verbose = FALSE` or `--quiet`.
