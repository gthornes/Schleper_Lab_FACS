library(openxlsx)

master_file <- "Loki_Cultures_Master_Reference.xlsx"
lookup_file <- "links_lookup.xlsx"
output_file <- "Loki_Cultures_Master_Reference.xlsx"

# ----------------------------
# Read lookup table
# ----------------------------

lookup <- read.xlsx(lookup_file)
names(lookup) <- trimws(names(lookup))

if (!all(c("Batch.File", "Cytogram.Link") %in% names(lookup))) {
  stop("links_lookup.xlsx must contain columns 'Batch.File' and 'Cytogram.Link'")
}

lookup$Batch.File <- trimws(as.character(lookup$Batch.File))
lookup$Cytogram.Link <- trimws(as.character(lookup$Cytogram.Link))

link_map <- setNames(lookup$Cytogram.Link, lookup$Batch.File)

# ----------------------------
# Load workbook
# ----------------------------

wb <- loadWorkbook(master_file)
sheets <- names(wb)

print(sheets)

group_border <- createStyle(
  border = "bottom",
  borderStyle = "thick")

for (sheet in sheets) {

  cat("\nProcessing sheet:", sheet, "\n")

  dat <- readWorkbook(wb, sheet = sheet)

  names(dat) <- trimws(names(dat))

  needed <- c(
    "Cytogram.Link",
    "Microscopy.Link",
    "Microscopy.URL",
    "Batch.File"
  )

  missing_cols <- setdiff(needed, names(dat))

  if (length(missing_cols) > 0) {
    cat("Skipping sheet. Missing columns:\n")
    print(missing_cols)
    next
  }

  # Find actual column positions from headers
  cyto_col <- which(names(dat) == "Cytogram.Link")
  micro_col <- which(names(dat) == "Microscopy.Link")
  micro_url_col <- which(names(dat) == "Microscopy.URL")
  batch_col <- which(names(dat) == "Batch.File")
  culture_col <- if ("Culture.ID" %in% names(dat)) which(names(dat) == "Culture.ID") else NA_integer_
  treatment_col <- if ("Treatment" %in% names(dat)) which(names(dat) == "Treatment") else NA_integer_

  blank_na_text <- function(x) {
    x <- trimws(as.character(x))
    x[is.na(x) | x == "NA"] <- ""
    x
  }

  dat$Batch.File <- blank_na_text(dat$Batch.File)
  dat$Microscopy.URL <- blank_na_text(dat$Microscopy.URL)
  if ("Culture.ID" %in% names(dat)) dat$Culture.ID <- blank_na_text(dat$Culture.ID)
  if ("Treatment" %in% names(dat)) dat$Treatment <- blank_na_text(dat$Treatment)

  for (i in seq_len(nrow(dat))) {

    row_num <- i + 1

    key <- dat$Batch.File[i]

    # --------------------------------
    # Cytogram link from lookup table
    # --------------------------------

    cyto_url <- ""

    if (!is.na(key) &&
        nzchar(key) &&
        key %in% names(link_map)) {

      cyto_url <- link_map[[key]]
    }

    if (!is.na(cyto_url) && nzchar(cyto_url)) {

      writeFormula(
        wb,
        sheet = sheet,
        x = sprintf(
          '=HYPERLINK("%s","Open Cytogram")',
          cyto_url
        ),
        startCol = cyto_col,
        startRow = row_num
      )

    } else {

      writeData(
        wb,
        sheet = sheet,
        x = "",
        startCol = cyto_col,
        startRow = row_num,
        colNames = FALSE
      )
    }

    micro_url <- dat$Microscopy.URL[i]

    if (!is.na(micro_url) && nzchar(micro_url)) {

      writeFormula(
        wb,
        sheet = sheet,
        x = sprintf(
          '=HYPERLINK("%s","Open Microscopy")',
          micro_url
        ),
        startCol = micro_col,
        startRow = row_num
      )

    } else {

      writeData(
        wb,
        sheet = sheet,
        x = "",
        startCol = micro_col,
        startRow = row_num,
        colNames = FALSE
      )
    }
  
  }

  # Hide Batch.File column
  setColWidths(
    wb,
    sheet = sheet,
    cols = batch_col,
    hidden = TRUE
  )

  setColWidths(
    wb,
    sheet = sheet,
    cols = micro_url_col,
    hidden = TRUE
  )

  if (!is.na(culture_col)) {
    for (i in seq_len(nrow(dat))) {
      if (!nzchar(dat$Culture.ID[i])) {
        writeData(
          wb,
          sheet = sheet,
          x = "",
          startCol = culture_col,
          startRow = i + 1,
          colNames = FALSE
        )
      }
    }
  }

  if (!is.na(treatment_col)) {
    for (i in seq_len(nrow(dat))) {
      if (!nzchar(dat$Treatment[i])) {
        writeData(
          wb,
          sheet = sheet,
          x = "",
          startCol = treatment_col,
          startRow = i + 1,
          colNames = FALSE
        )
      }
    }
  }

  cat("Finished", sheet, "link population.\n")

    if ("Culture.ID" %in% names(dat)) {

      culture_ids <- trimws(as.character(dat$Culture.ID))
      culture_ids[is.na(culture_ids)] <- ""

      filled_ids <- culture_ids
      last_seen_id <- ""

      for (i in seq_along(filled_ids)) {
        if (nzchar(filled_ids[i])) {
          last_seen_id <- filled_ids[i]
        } else {
          filled_ids[i] <- last_seen_id
        }
      }

      for (i in 1:(nrow(dat) - 1)) {

          current_id <- filled_ids[i]
          next_id    <- filled_ids[i + 1]

          if (nzchar(current_id) && nzchar(next_id) && !identical(current_id, next_id)) {

          addStyle(
            wb,
            sheet = sheet,
            style = group_border,
            rows = i + 1,                # +1 because row 1 is headers
            cols = 1:ncol(dat),
            gridExpand = TRUE,
            stack = TRUE
          )
          }
      }
      }

    if ("Culture.ID" %in% names(dat)) {

    culture_col <- which(names(dat) == "Culture.ID")

    ids <- trimws(as.character(dat$Culture.ID))
    ids[is.na(ids)] <- ""

    for (i in 2:length(ids)) {

      if (ids[i] != "" &&
          ids[i] == ids[i - 1]) {

        writeData(
          wb,
          sheet = sheet,
          x = "",
          startCol = culture_col,
          startRow = i + 1,
          colNames = FALSE
        )
      }
    }
  }
}

saveWorkbook(
  wb,
  output_file,
  overwrite = TRUE
)

cat(
  "\nLinks populated successfully:\n",
  output_file,
  "\n"
)