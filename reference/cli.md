# Convert an XLSForm from the terminal

Runs
[`convertxlsform()`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)
with options taken from the command line. Use it from a shell without
opening R:

## Usage

``` r
cli(args = commandArgs(trailingOnly = TRUE))
```

## Arguments

- args:

  Character vector of arguments. Defaults to the trailing command-line
  arguments.

## Value

The exit status invisibly: 0 on success, 1 on a usage error or a failed
conversion. The function does not quit R.

## Details

    Rscript -e 'convertxlsform::cli()' form.xlsx --format pdf --lang bn

Put the XLSForm path and the options after the quoted expression. If a
script or CI job needs the exit status, use
`Rscript -e 'quit(status = convertxlsform::cli())' form.xlsx`. The plain
form above also works for normal use.

This function is a terminal interface. Its help and version text go to
standard output, and usage errors go to standard error, because that
output is its purpose. Progress messages are sent with
[`message()`](https://rdrr.io/r/base/message.html) and `--quiet`
suppresses them (the same as `verbose = FALSE` in
[`convertxlsform()`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)).

## Usage

`cli(args)`: `form.xlsx [options]`. Options accept `--name value` and
`--name=value`.

- `-o`, `--output PATH`: output file or folder (default: next to the
  form).

- `-f`, `--format docx|pdf|both`: output format (default: from the
  extension of `--output`, else `docx`).

- `-l`, `--lang LANG`: language code, name or header suffix.

- `--flat-numbers`: number questions 1, 2, 3 instead of 1.1, 1.2.

- `--no-relevants`: hide skip logic.

- `--no-constraints`: hide constraints.

- `--no-hints`: hide hints.

- `--no-choice-codes`: hide choice codes.

- `--show-names`: show variable names.

- `--include-hidden`: include calculate, start, end and similar rows.

- `--media-root DIR`: folder for media and external files (default: the
  folder of the form).

- `--title TEXT`: form title.

- `--page-size A4|letter` (default `A4`).

- `--landscape`: landscape orientation.

- `--margins IN`: margins in inches (default 1).

- `--font NAME` (default `Arial`) and `--font-size PT` (default 11).

- `--template FILE`: reference `.docx` for styles.

- `--max-choices N`: longest choice list printed in full (default 30).

- `--repeat-copies N`: copies of each repeat group (default 1).

- `--overwrite`: replace existing output.

- `--quiet`: no progress messages.

- `-h`, `--help`: show the help and stop.

- `--version`: show the package version and stop.

## See also

[`convertxlsform()`](https://masud90.github.io/convertxlsform/reference/convertxlsform.md)

## Examples

``` r
path <- system.file("extdata", "multilang", "multilang.xlsx",
                    package = "convertxlsform")
cli(c(path, "-o", tempdir(), "--overwrite", "--quiet"))
```
