# Generates inst/templates/convertxlsform.docx (reference Word template).
# Run from the project root: Rscript dev/make_template.R
library(officer)

out <- file.path("inst", "templates", "convertxlsform.docx")
dir.create(dirname(out), recursive = TRUE, showWarnings = FALSE)

x <- read_docx()
x <- body_set_default_section(x, prop_section(
  page_size = page_size(orient = "portrait", width = 8.27, height = 11.69),
  page_margins = page_mar(top = 1, bottom = 1, left = 1, right = 1,
                          header = 0.5, footer = 0.5, gutter = 0),
  type = "continuous"
))

font <- "Arial"
tx <- function(size = 11, bold = FALSE, italic = FALSE, color = "black") {
  fp_text_lite(font.family = font, font.size = size, bold = bold,
               italic = italic, color = color)
}
pp <- function(...) fp_par(..., line_spacing = 1.1)

para <- function(x, id, name, fp_p, fp_t, base = "Normal") {
  docx_set_paragraph_style(x, style_id = id, style_name = name,
                           base_on = base, fp_p = fp_p, fp_t = fp_t)
}

# Existing styles keep their ids (ids in officer's default template may be localized).
ids <- styles_info(x)
id_of <- function(nm) ids$style_id[ids$style_name == nm & ids$style_type == "paragraph"][1]

x <- para(x, "Title", "Title",
          pp(text.align = "center", padding.bottom = 8, padding.top = 0),
          tx(16, bold = TRUE))
hd <- list(c(13, 10), c(12, 8), c(11, 6))
for (i in 1:3) {
  nm <- paste("heading", i)
  x <- para(x, id_of(nm), nm,
            pp(keep_with_next = TRUE, padding.top = hd[[i]][2], padding.bottom = 4,
               border.bottom = fp_border(width = 0, style = "none")),
            tx(hd[[i]][1], bold = TRUE, color = "#1F3864"))
}
x <- para(x, "cxfQuestion", "cxf Question",
          pp(keep_with_next = TRUE, padding.top = 8, padding.bottom = 2), tx(11))
x <- para(x, "cxfHint", "cxf Hint",
          pp(keep_with_next = TRUE, padding.bottom = 2),
          tx(10, italic = TRUE, color = "#595959"))
x <- para(x, "cxfLogic", "cxf Logic",
          pp(keep_with_next = TRUE, padding.bottom = 2), tx(10, color = "#7F6000"))
x <- para(x, "cxfChoice", "cxf Choice",
          pp(padding.left = 0.3 * 72, padding.bottom = 2), tx(11))
x <- para(x, "cxfNote", "cxf Note",
          pp(shading.color = "#F2F2F2", padding.top = 4, padding.bottom = 4),
          tx(11, italic = TRUE))
x <- para(x, "cxfBanner", "cxf Banner",
          pp(shading.color = "#DEEAF6", keep_with_next = TRUE,
             padding.top = 6, padding.bottom = 4),
          tx(10, bold = TRUE, color = "#1F3864"))

x <- set_doc_properties(x, title = "convertxlsform template")
print(x, target = out)

# Normal: docx_set_paragraph_style() would make it basedOn itself and drop its
# default flag, so patch the style in place in the saved package.
tmp <- tempfile("cxf"); dir.create(tmp)
utils::unzip(out, exdir = tmp)
sf <- file.path(tmp, "word", "styles.xml")
doc <- xml2::read_xml(sf)
nrm <- xml2::xml_find_first(doc, "//w:style[@w:styleId='Normal']")
for (nd in xml2::xml_children(nrm)) {
  if (xml2::xml_name(nd) %in% c("pPr", "rPr", "basedOn")) xml2::xml_remove(nd)
}
xml2::xml_attr(nrm, "w:customStyle") <- NULL
xml2::xml_attr(nrm, "w:default") <- "1"
w <- 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
xml2::xml_add_child(nrm, xml2::read_xml(paste0("<w:pPr ", w,
  '><w:spacing w:before="0" w:after="80" w:line="264" w:lineRule="auto"/></w:pPr>')))
xml2::xml_add_child(nrm, xml2::read_xml(paste0("<w:rPr ", w,
  '><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>',
  '<w:color w:val="000000"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>')))
xml2::write_xml(doc, sf)
unlink(out)
owd <- setwd(tmp)
zip::zip(file.path(owd, out), list.files(".", all.files = TRUE, recursive = TRUE),
         mode = "mirror")
setwd(owd)
message("Wrote ", out, " (", file.size(out), " bytes)")
