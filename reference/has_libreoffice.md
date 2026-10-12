# Is LibreOffice available?

Reports whether a LibreOffice binary can be found, which is needed to
convert Word documents to PDF with
[`docx_to_pdf()`](https://masud90.github.io/convertxlsform/reference/docx_to_pdf.md).
The search checks the option `convertxlsform.soffice`, the environment
variable `CONVERTXLSFORM_SOFFICE`, the `PATH`, and the default install
locations for the operating system.

## Usage

``` r
has_libreoffice()
```

## Value

`TRUE` or `FALSE`.

## Examples

``` r
has_libreoffice()
#> [1] FALSE
```
