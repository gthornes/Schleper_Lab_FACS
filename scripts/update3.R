#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(openxlsx)
})

options("openxlsx.dateFormat" = "dd.mm.yyyy")

args <- commandArgs(trailingOnly = TRUE)
input_path        <- if (length(args) >= 1) args[1] else "."
flow_lookup_path  <- if (length(args) >= 2) args[2] else "flow_rate_lookup.csv"
out_path          <- if (length(args) >= 3) args[3] else "Loki_Cultures_Master_Reference.xlsx"

if (!file.exists(flow_lookup_path)) stop("Missing flow lookup file: ", flow_lookup_path)

find_csv_files <- function(path) {
  if (file.exists(path) && grepl("\\.csv$", path, ignore.case = TRUE)) return(normalizePath(path))
  if (!dir.exists(path)) stop("Input path not found: ", path)
  files <- list.files(path, pattern = "Batch_Analysis_.*\\.csv$", recursive = TRUE, full.names = TRUE)
  if (length(files) == 0) stop("No Batch_Analysis_*.csv files found in ", path)
  normalizePath(files)
}

get_col <- function(df, candidates, required = TRUE) {
  norm <- function(x) gsub("[^a-z0-9]+", "", tolower(x))
  name_norm <- norm(names(df))
  for (cand in candidates) {
    idx <- match(norm(cand), name_norm)
    if (!is.na(idx)) return(names(df)[idx])
  }
  if (required) stop("Missing column: ", paste(candidates, collapse = " or "))
  return(NULL)
}

safe_div <- function(num, den) ifelse(is.na(den) | den == 0, NA_real_, num / den)

as_numeric_safe <- function(x) {
  if (is.numeric(x)) return(x)
  x_chr <- trimws(as.character(x))
  x_chr[x_chr == ""] <- NA_character_
  suppressWarnings(as.numeric(gsub(",", ".", x_chr, fixed = TRUE)))
}

culture_sort_key <- function(x) {
  x_chr <- trimws(as.character(x))
  x_num <- suppressWarnings(as.numeric(x_chr))
  ifelse(is.na(x_num), x_chr, sprintf("%020.0f", x_num))
}

norm_key <- function(x) gsub("[^a-z0-9]+", "", tolower(trimws(as.character(x))))

norm_date_chr <- function(x) {
  if (inherits(x, "Date")) return(format(x, "%d.%m.%Y"))
  d <- suppressWarnings(as.Date(as.character(x), format = "%Y-%m-%d"))
  d2 <- suppressWarnings(as.Date(as.character(x), format = "%d.%m.%Y"))
  out <- ifelse(!is.na(d), format(d, "%d.%m.%Y"), as.character(x))
  out <- ifelse(is.na(d) & !is.na(d2), format(d2, "%d.%m.%Y"), out)
  out
}

extract_ddmm <- function(x) {
  x <- as.character(x)
  m <- regexpr("\\d{4}$", x)
  candidate <- ifelse(m == -1, NA_character_, regmatches(x, m))
  day   <- suppressWarnings(as.integer(substr(candidate, 1, 2)))
  month <- suppressWarnings(as.integer(substr(candidate, 3, 4)))
  valid <- !is.na(day) & !is.na(month) & day >= 1 & day <= 31 & month >= 1 & month <= 12
  ifelse(valid, candidate, NA_character_)
}

make_sample_date <- function(tube, record_date) {
  tube <- trimws(as.character(tube))
  record_date <- trimws(as.character(record_date))

  parsed_dt <- suppressWarnings(as.POSIXct(record_date, format = "%b %d, %Y %I:%M:%S %p", tz = "UTC"))
  if (is.na(parsed_dt)) {
    parsed_dt <- suppressWarnings(as.POSIXct(record_date, format = "%B %d, %Y %I:%M:%S %p", tz = "UTC"))
  }
  if (is.na(parsed_dt)) {
    return(as.Date(NA))
  }

  year_val <- format(parsed_dt, "%Y")

  m <- regexec(".*_(\\d{4})$", tube, perl = TRUE)
  hit <- regmatches(tube, m)[[1]]
  if (length(hit) < 2) {
    return(as.Date(NA))
  }

  ddmm <- hit[2]
  day <- suppressWarnings(as.integer(substr(ddmm, 1, 2)))
  mon <- suppressWarnings(as.integer(substr(ddmm, 3, 4)))

  if (is.na(day) || is.na(mon) || day < 1 || day > 31 || mon < 1 || mon > 12) {
    return(as.Date(NA))
  }

  as.Date(sprintf("%s-%02d-%02d", year_val, mon, day))
}

normalize_existing <- function(existing, headers) {
  if (nrow(existing) == 0) {
    return(setNames(data.frame(matrix(ncol = length(headers), nrow = 0), check.names = FALSE), headers))
  }

  clean_existing <- gsub("[^a-zA-Z0-9]+", ".", names(existing))
  clean_headers  <- gsub("[^a-zA-Z0-9]+", ".", headers)

  for (i in seq_along(headers)) {
    hit <- match(clean_headers[i], clean_existing)
    if (!is.na(hit)) names(existing)[hit] <- headers[i]
  }

  missing_cols <- setdiff(headers, names(existing))
  for (col in missing_cols) existing[[col]] <- NA
  existing <- existing[, headers, drop = FALSE]

  fix_date_col <- function(x) {
    if (inherits(x, "Date")) return(x)

    x_chr <- as.character(x)
    from_dmy <- suppressWarnings(as.Date(x_chr, format = "%d.%m.%Y"))
    from_iso <- suppressWarnings(as.Date(x_chr, format = "%Y-%m-%d"))

    num <- suppressWarnings(as.numeric(x_chr))
    from_serial <- suppressWarnings(as.Date(num, origin = "1899-12-30"))
    year_serial <- suppressWarnings(as.integer(format(from_serial, "%Y")))
    use_serial <- !is.na(from_serial) & !is.na(year_serial) & year_serial >= 1990 & year_serial <= 2100

    result <- as.Date(rep(NA, length(x_chr)))
    result[!is.na(from_dmy)] <- from_dmy[!is.na(from_dmy)]
    result[is.na(result) & !is.na(from_iso)] <- from_iso[is.na(result) & !is.na(from_iso)]
    result[is.na(result) & use_serial] <- from_serial[is.na(result) & use_serial]
    result
  }

  if ("Sample Date" %in% names(existing)) {
    existing$`Sample Date` <- fix_date_col(existing$`Sample Date`)
  }
  if ("Inoculation Date" %in% names(existing)) {
    existing$`Inoculation Date` <- fix_date_col(existing$`Inoculation Date`)
  }

  existing
}

append_by_culture <- function(existing_df, new_df) {

  if (nrow(existing_df) == 0)
    return(new_df)

  make_key <- function(df) {
    paste(df$`Culture ID`, norm_date_chr(df$`Sample Date`), sep = "|")
  }

  sort_block_by_date <- function(df) {
    if (nrow(df) <= 1) return(df)
    sample_dates <- suppressWarnings(as.Date(norm_date_chr(df$`Sample Date`), format = "%d.%m.%Y"))
    date_key <- ifelse(is.na(sample_dates), Inf, as.numeric(sample_dates))
    df[order(date_key, seq_len(nrow(df))), , drop = FALSE]
  }

  fill_inoculation_from_block <- function(block, add_rows) {
    if (!"Inoculation Date" %in% names(block) || !"Inoculation Date" %in% names(add_rows)) {
      return(add_rows)
    }

    existing_inoc <- block$`Inoculation Date`
    existing_inoc_chr <- trimws(as.character(existing_inoc))
    valid_existing <- !is.na(existing_inoc) & nzchar(existing_inoc_chr) & existing_inoc_chr != "NA"

    if (!any(valid_existing)) {
      return(add_rows)
    }

    inoc_value <- existing_inoc[which(valid_existing)[1]]
    add_inoc_chr <- trimws(as.character(add_rows$`Inoculation Date`))
    needs_fill <- is.na(add_rows$`Inoculation Date`) | !nzchar(add_inoc_chr) | add_inoc_chr == "NA"
    add_rows$`Inoculation Date`[needs_fill] <- inoc_value
    add_rows
  }

  out <- existing_df

  for (key in unique(make_key(new_df))) {

    add_rows <- new_df[make_key(new_df) == key, , drop = FALSE]
    culture_value <- add_rows$`Culture ID`[1]
    culture_idx <- which(out$`Culture ID` == culture_value)

    if (length(culture_idx) == 0) {
      out <- rbind(out, add_rows)
      next
    }

    block <- out[culture_idx, , drop = FALSE]
    add_rows <- fill_inoculation_from_block(block, add_rows)

    idx <- which(make_key(block) == key)
    if (length(idx) > 0) {
      block <- block[-idx, , drop = FALSE]
    }

    block <- rbind(block, add_rows)
    block <- sort_block_by_date(block)

    before_idx <- min(culture_idx)
    after_idx <- max(culture_idx)

    before_block <- if (before_idx > 1) out[seq_len(before_idx - 1), , drop = FALSE] else out[0, , drop = FALSE]
    after_block <- if (after_idx < nrow(out)) out[(after_idx + 1):nrow(out), , drop = FALSE] else out[0, , drop = FALSE]

    out <- rbind(before_block, block, after_block)
  }

  rownames(out) <- NULL

  if (nrow(out) == 0) {
    return(out)
  }

  culture_key <- culture_sort_key(out$`Culture ID`)
  out[order(culture_key, seq_along(culture_key), na.last = TRUE), , drop = FALSE]
}

flow_lookup <- read.csv(flow_lookup_path, stringsAsFactors = FALSE)

if (!"lookup_key" %in% names(flow_lookup)) {
  if (!all(c("cell_line", "sample") %in% names(flow_lookup))) {
    stop("flow_lookup.csv must contain either lookup_key, or both cell_line and sample columns")
  }
  flow_lookup$lookup_key <- paste(norm_key(flow_lookup$cell_line), norm_key(flow_lookup$sample), sep = "|")
} else {
  flow_lookup$lookup_key <- norm_key(flow_lookup$lookup_key)
}

csv_files <- find_csv_files(input_path)

headers_b35 <- c(
  "Culture ID", "Sample Date", "Inoculation Date", "Age (days)",
  "Loki mL-1", "Loki %", "DSV/Spiro mL-1", "DSV/Spiro %",
  "Cytogram Link", "Microscopy Link", "Microscopy URL", "Treatment", "Batch File"
)

headers_b36 <- c(
  "Culture ID", "Sample Date", "Inoculation Date", "Age (days)",
  "Loki mL-1", "Loki %", "DSV/Spiro mL-1", "DSV/Spiro %",
  "Izemo mL-1", "Izemo %",
  "Cytogram Link", "Microscopy Link", "Microscopy URL", "Treatment", "Batch File"
)

all_new_rows <- list()

for (csv_path in csv_files) {
  batch_pdf_file <- sub("\\.csv$", ".pdf", basename(csv_path), ignore.case = TRUE)
  df <- read.csv(csv_path, stringsAsFactors = FALSE, check.names = FALSE)

  col_tube   <- get_col(df, c("Tube Name"))
  col_record <- get_col(df, c("Record Date"))
  col_all    <- get_col(df, c("All Events #Events"))
  col_noise  <- get_col(df, c("noise #Events"))
  col_loki   <- get_col(df, c("Loki #Events", "loki #Events"))
  col_dsv    <- get_col(df, c("DSV #Events"))
  col_izemo  <- get_col(df, c("Izemo #Events"), required = FALSE)

  tube <- as.character(df[[col_tube]])
  record_date <- as.character(df[[col_record]])

  parts <- strsplit(tube, "_")
  strain <- vapply(parts, function(x) if (length(x) >= 1) x[[1]] else NA_character_, character(1))
  culture_id <- vapply(parts, function(x) if (length(x) >= 2) x[[2]] else NA_character_, character(1))
  sample_suffix <- vapply(parts, function(x) if (length(x) >= 3) paste(x[3:length(x)], collapse = "_") else NA_character_, character(1))

  sample_name <- ifelse(is.na(sample_suffix) | sample_suffix == "", culture_id, paste(culture_id, sample_suffix, sep = "_"))
  lookup_key  <- paste(norm_key(strain), norm_key(sample_name), sep = "|")
  match_idx   <- match(lookup_key, flow_lookup$lookup_key)

  dilution_sample  <- as_numeric_safe(flow_lookup$dilution[match_idx])
  volume_passed_ul <- as_numeric_safe(flow_lookup$volume_passed_ul[match_idx])
  dilution_factor  <- 1000 / (1000 + dilution_sample)
  sample_ul <- 500
  sybr_ul <- 5

  scaling <- safe_div(1000, volume_passed_ul) / dilution_factor * dilution_sample * ((sample_ul + sybr_ul) / sample_ul)

  all_ml   <- as_numeric_safe(df[[col_all]])   * scaling
  noise_ml <- as_numeric_safe(df[[col_noise]]) * scaling
  loki_ml  <- as_numeric_safe(df[[col_loki]])  * scaling
  dsv_ml   <- as_numeric_safe(df[[col_dsv]])   * scaling
  denom    <- all_ml - noise_ml

  loki_pct <- safe_div(loki_ml, denom) * 100
  dsv_pct  <- safe_div(dsv_ml, denom) * 100

  izemo_ml  <- if (!is.null(col_izemo)) as_numeric_safe(df[[col_izemo]]) * scaling else rep(NA_real_, nrow(df))
  izemo_pct <- safe_div(izemo_ml, denom) * 100

  sample_date <- vapply(
    seq_along(tube),
    function(i) as.character(make_sample_date(tube[i], record_date[i])),
    character(1)
  )
  sample_date <- as.Date(sample_date)

  new_rows <- data.frame(
    `Culture ID`       = culture_id,
    `Sample Date`      = sample_date,
    `Inoculation Date` = as.Date(rep(NA, length(culture_id)), origin = "1970-01-01"),
    `Age (days)`       = rep(NA_real_, length(culture_id)),
    `Loki mL-1`        = loki_ml,
    `Loki %`           = loki_pct,
    `DSV/Spiro mL-1`   = dsv_ml,
    `DSV/Spiro %`      = dsv_pct,
    `Izemo mL-1`       = izemo_ml,
    `Izemo %`          = izemo_pct,
    `Cytogram Link`    = "",
    `Microscopy Link`  = "",
    `Microscopy URL`   = "",
    `Treatment`        = "",
    `Batch File`       = batch_pdf_file,
    sheet              = strain,
    stringsAsFactors   = FALSE,
    check.names        = FALSE
  )

  all_new_rows[[csv_path]] <- new_rows
}

combined <- do.call(rbind, all_new_rows)
if (nrow(combined) == 0) stop("No rows were generated from the input files.")

style_header  <- createStyle(textDecoration = "bold", fgFill = "#2E4057", fontColour = "#FFFFFF", halign = "center", valign = "center", border = "Bottom", borderColour = "#1A2A3A", borderStyle = "medium")
style_culture <- createStyle(textDecoration = "bold", fgFill = "#E8F0F7")
style_date    <- createStyle(numFmt = "dd.mm.yyyy", fgFill = "#F5F5F5", halign = "center")
style_age     <- createStyle(fgFill = "#EEE8F4", halign = "center", numFmt = "0")
style_loki    <- createStyle(fgFill = "#D6EAD6")
style_dsv     <- createStyle(fgFill = "#FDE8D8")
style_izemo   <- createStyle(fgFill = "#FFF3CD")
style_link    <- createStyle(fgFill = "#F0F0F0", fontColour = "#2D6DB5", textDecoration = "underline")
style_alt     <- createStyle(fgFill = "#FAFAFA")
style_sci     <- createStyle(numFmt = "0.00E+00")
style_pct     <- createStyle(numFmt = "0.0\"%\"")

wb <- if (file.exists(out_path)) loadWorkbook(out_path) else createWorkbook()

sheet_config <- list(
  B35 = list(
    headers = headers_b35,
    link_col = 9,
    n_cols = 13,
    col_widths = c(14, 14, 14, 11, 14, 10, 16, 12, 16, 16, 0.1, 16, 0.1)),
  B36 = list(
    headers = headers_b36,
    link_col = 11,
    n_cols = 15,
    col_widths = c(14, 14, 14, 11, 14, 10, 16, 12, 14, 10, 16, 16, 0.1, 16, 0.1)))

for (sheet in names(sheet_config)) {
  cfg     <- sheet_config[[sheet]]
  hdrs    <- cfg$headers
  lnk_col <- cfg$link_col
  nc      <- cfg$n_cols

  sheet_rows <- combined[combined$sheet == sheet, c(hdrs, "Batch File"), drop = FALSE]

  existing <- if (file.exists(out_path)) {
    tryCatch(read.xlsx(out_path, sheet = sheet, detectDates = TRUE), error = function(e) data.frame())
  } else {
    data.frame()
  }

  existing <- normalize_existing(existing, hdrs)

  if ("Culture ID" %in% names(existing)) {

    last_id <- NA_character_

    for (i in seq_len(nrow(existing))) {

      current <- trimws(as.character(existing$`Culture ID`[i]))

      if (!is.na(current) && nzchar(current)) {
        last_id <- current
      } else {
        existing$`Culture ID`[i] <- last_id
      }
    }
  }
  updated  <- append_by_culture(existing, sheet_rows[, hdrs, drop = FALSE])

  blank_na_text <- function(x) {
    if (is.character(x)) {
      x[is.na(x) | trimws(x) == "NA"] <- ""
    }
    x
  }

  updated[] <- lapply(updated, blank_na_text)

  if (sheet %in% names(wb)) removeWorksheet(wb, sheet)
  addWorksheet(wb, sheet, gridLines = FALSE)

  writeData(wb, sheet, updated, startRow = 1, headerStyle = style_header, keepNA = FALSE)
  setColWidths(wb, sheet, cols = seq_len(nc), widths = cfg$col_widths)

  n_rows <- nrow(updated)
  if (n_rows > 0) {
    data_rows <- 2:(n_rows + 1)
    even_rows <- data_rows[data_rows %% 2 == 0]

    if (length(even_rows) > 0) {
      addStyle(wb, sheet, style_alt, rows = even_rows, cols = seq_len(nc), gridExpand = TRUE, stack = TRUE)
    }

    addStyle(wb, sheet, style_culture, rows = data_rows, cols = 1, stack = TRUE)
    addStyle(wb, sheet, style_date,    rows = data_rows, cols = 2:3, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, style_age,     rows = data_rows, cols = 4, stack = TRUE)
    addStyle(wb, sheet, style_loki,    rows = data_rows, cols = 5:6, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, style_dsv,     rows = data_rows, cols = 7:8, gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, style_sci,     rows = data_rows, cols = c(5, 7), gridExpand = TRUE, stack = TRUE)
    addStyle(wb, sheet, style_pct,     rows = data_rows, cols = c(6, 8), gridExpand = TRUE, stack = TRUE)

    if (sheet == "B36") {
      addStyle(wb, sheet, style_izemo, rows = data_rows, cols = 9:10, gridExpand = TRUE, stack = TRUE)
      addStyle(wb, sheet, style_sci,   rows = data_rows, cols = 9, stack = TRUE)
      addStyle(wb, sheet, style_pct,   rows = data_rows, cols = 10, stack = TRUE)
    }

    addStyle(wb, sheet, style_link, rows = data_rows, cols = lnk_col:(lnk_col + 1), gridExpand = TRUE, stack = TRUE)

    age_formulas <- sprintf('IF(OR(C%d="",B%d=""),"",B%d-C%d)', data_rows, data_rows, data_rows, data_rows)
    writeFormula(wb, sheet = sheet, x = age_formulas, startCol = 4, startRow = 2)

    batch_col <- which(names(updated) == "Batch File")
    micro_url_col <- which(names(updated) == "Microscopy URL")

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
    freezePane(wb, sheet, firstRow = TRUE)
  }
}

saveWorkbook(wb, out_path, overwrite = TRUE)
cat("Updated mastersheet:", out_path, "\n")