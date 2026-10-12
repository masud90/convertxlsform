# CRAN submission guide for convertxlsform 0.3.0

Owner's step-by-step guide. Nothing has been submitted to CRAN.

## Status

Local checks verified on 2026-10-12, on macOS 27.0.1 (aarch64), R 4.6.0:

- [x] `R CMD build` with vignettes (pandoc 3.10 from RStudio via `RSTUDIO_PANDOC`).
- [x] `R CMD check --as-cran convertxlsform_0.3.0.tar.gz` with the PDF manual: 0 errors, 0 warnings, 2 notes (New submission; local HTML Tidy too old). Details in `cran-comments.md`.
- [x] `testthat::test_local()` with `NOT_CRAN=true`: 0 failures, 0 warnings, 1 skip (doconv/Word test), 1058 passes.
- [x] `spelling::spell_check_package()`: no errors (`inst/WORDLIST` updated).
- [x] `urlchecker::url_check()`: all URLs correct.
- [x] R-hub v2 workflow is on `main` (`.github/workflows/rhub.yaml`).
- [x] PR #1 merged into `main`. GitHub Actions R-CMD-check was green on macOS, Windows, and Ubuntu release and oldrel-1.
- [x] `v0.3.0` tagged and released on GitHub: <https://github.com/masud90/convertxlsform/releases/tag/v0.3.0>.
- [x] Win-builder R-devel (first run): 2 NOTEs. (1) CRAN incoming: new submission plus a possibly misspelled word, 'docx' in DESCRIPTION; fixed by quoting it as `'.docx'`. (2) Examples with elapsed time over 10 s (`cli` 12.90 s, `convertxlsform` 11.19 s); fixed by pointing those examples at a small form (`inst/extdata/minimal/minimal.xlsx`) and moving the larger ones into `\donttest{}`. The tests also took 308 s there; the heavy rendering tests are now `skip_on_cran()`. Re-run on 2026-10-12 after the fixes: 1 NOTE (new submission). Examples OK, tests 80 s.
- [x] Win-builder R-release (after the fixes, 2026-10-12): 1 NOTE (new submission). Examples OK, tests 81 s.
- [x] GitHub Actions on `main` (f81b5b6): macOS, Windows, Ubuntu devel, release and oldrel-1 all passed.
- [ ] Not yet run: mac builder, R-hub. No results exist for these.

The tarball from that local check was built in a temporary folder and is gone. Rebuild it from the current `main` before every check or submission (see "How to submit").

## Before submitting

- [x] Review and commit the session changes, and push them.
- [x] CI green on PR #1 (macOS, Windows, Ubuntu release and oldrel-1).
- [x] Merge PR #1 into `main`. The R-hub workflow is on the default branch, as `rhub_check()` requires.
- [x] Tag `v0.3.0` on `main`.
- [x] Win-builder R-devel re-run and R-release: both 1 NOTE (new submission). `devtools` is not installed, so either install it (`install.packages("devtools")`) and run:
  ```r
  devtools::check_win_devel()
  devtools::check_win_release()
  ```
  or upload the tarball at <https://win-builder.r-project.org/upload.aspx> for R-devel and R-release. Results arrive by email to the maintainer address.
- [ ] Mac builder: upload the tarball at <https://mac.r-project.org/macbuilder/submit.html>.
- [ ] R-hub v2: `rhub::rhub_doctor()`, then
  ```r
  rhub::rhub_check(platforms = c("linux", "macos-arm64", "windows", "ubuntu-next", "ubuntu-release"))
  ```
  Run `rhub::rhub_platforms()` first and use platform names that it lists.
- [ ] Copy the results into `cran-comments.md`, replacing each `[to fill ...]` placeholder with the actual outcome. Do not leave placeholders in the submitted file.
- [ ] Rebuild the tarball (below) if any file changed after the last build.

## How to submit

1. Build the tarball with vignettes from a temporary directory. pandoc is not on the PATH here, so point to the one in RStudio:
   ```sh
   export RSTUDIO_PANDOC=/Applications/RStudio.app/Contents/Resources/app/quarto/bin/tools/aarch64
   cd "$(mktemp -d)"
   R CMD build /Users/masud/Dev/convertxlsform
   R CMD check --as-cran convertxlsform_0.3.0.tar.gz
   ```
   For the PDF manual, `~/Library/TinyTeX/bin/universal-darwin` must be on the PATH.
2. Open <https://cran.r-project.org/submit.html>, upload the tarball, paste the comments from `cran-comments.md`, and submit.
3. CRAN sends a confirmation email to the maintainer address `write2masud+convertxlsform@gmail.com` (it reaches write2masud@gmail.com). The submission is not complete until you open the link in that email and confirm.
4. Alternative: `devtools::release()` runs the checklist and submits.
5. Reply to CRAN's emails by replying to the same thread. Fix, bump nothing unless asked, and resubmit with a short note on what changed.

## What CRAN reviewers commonly ask for on first submissions

| Request | Where the package stands |
|---|---|
| `\value` in every exported Rd | Done for all six exports, including `cli()` and `has_libreoffice()`. |
| Software names in single quotes in Title and Description | Done: 'XLSForm', 'Word', 'LibreOffice', 'ODK', 'KoboToolbox', 'Open Data Kit'. |
| Web reference for the method or format in Description | Done: `<https://xlsform.org>`. |
| Copyright holder in Authors@R | Done: role "cph" added. |
| Title in title case, under 65 characters, no period | Done (49 characters). |
| No "This package" or package name at the start of Description | Done. |
| No `\dontrun{}` | None used. `\donttest{}` only for the slower second example of `convertxlsform()`. |
| Interactive-only code guarded | PDF examples use `if (interactive() && has_libreoffice())`. |
| No `print()`/`cat()` for information | None outside `cli()`, whose help and version output is its purpose. This is stated in the `cli()` documentation. Progress uses cli messages and `verbose = FALSE` or `--quiet` silences them. |
| Options, `par()`, working directory restored | The package changes none of them. Temporary files are removed with `on.exit()`. |
| No writing to the user's filespace in examples, tests, vignette | All use `tempdir()`. Default output location: see Risks. |
| Examples run in under 5 s | `R CMD check` examples (including `--run-donttest`) passed. The check log does not report per-example times, so run `devtools::run_examples()` if CRAN questions this. |
| Spelling and URLs | Both checks are clean. |
| Vignette builds | Built locally with pandoc 3.10. |

## Risks specific to this package

1. **Output location default.** With `output = NULL`, `convertxlsform()` writes next to the input file. Reviewers sometimes flag "functions should not write by default to the user's filespace". The behaviour was not changed. The justification is in `cran-comments.md`: the location derives from the path the user supplied, it is reported in a message, `output` lets the user choose, and existing files are protected unless `overwrite = TRUE`. **Fallback if a reviewer insists:** make `output` default to `tempdir()` in `.cxf_targets()` in `R/convertxlsform.R`, update the docs, README, vignette, CLI help and tests that assert the default (`test-convertxlsform.R`, `test-cli.R`), regenerate docs, rerun the checks, and note it in `NEWS.md`.
2. **LibreOffice in SystemRequirements.** It is optional and only used for PDF. If asked, point to the guards: examples run only when interactive, tests skip without it and on CRAN. The field text is `'LibreOffice' (optional, for PDF output)`.
3. **Test length.** The tests take about 49 s locally under `--as-cran`. CRAN's limit is much higher, but if a platform is slow, add `skip_on_cran()` to the heaviest render tests.
4. **pandoc.** Not on the PATH on this machine. The vignette needs it to build, so use `RSTUDIO_PANDOC` as above. CRAN machines have pandoc.
5. **Dependencies.** Imports include `flextable` and `officer`, which pull in many packages. This is usually accepted, but a reviewer may comment on the installed size.

## After acceptance

- [ ] Change the README install lines to `install.packages("convertxlsform")`, keep the GitHub line for the development version.
- [ ] Add the CRAN badge to the README (`usethis::use_cran_badge()`, or the shields.io CRAN version badge).
- [ ] Check the NEWS heading is `# convertxlsform 0.3.0`.
- [ ] Bump `DESCRIPTION` to `0.3.0.9000` and add a `# convertxlsform (development version)` heading at the top of `NEWS.md`.
- [ ] Confirm the pkgdown site deploys (the workflow runs on push to `main`) and that the GitHub Pages setting uses the `gh-pages` branch.
- [ ] Update `CLAUDE.md` status.

## Known issue deferred to 0.4.0

PDF output can show blank text for Bengali and other complex scripts when LibreOffice cannot find a suitable font. The Word document is correct. It is listed under "Known issues" in `NEWS.md` and is not a CRAN blocker.
