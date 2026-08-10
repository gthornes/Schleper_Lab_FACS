#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  library(readxl)
})

args <- commandArgs(trailingOnly = TRUE)
fac_path <- if (length(args) >= 1) args[1] else "FACS_0603.xlsx"
out_csv  <- if (length(args) >= 2) args[2] else file.path("output", "flow_rate_lookup_0603.csv")

out_dir <- dirname(out_csv)
if (!dir.exists(out_dir)) {
  dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
}

if (!file.exists(fac_path)) {
  stop("Missing input file: ", fac_path)
}

raw <- read_excel(fac_path, col_names = FALSE)
flow_row <- which(raw[[1]] == "Flow Rate")[1]
sample_header_row <- which(raw[[1]] == "Sample Number")[1]

if (is.na(flow_row) || is.na(sample_header_row)) {
  stop("Could not find calibration or sample header rows in ", fac_path)
}

calib <- raw[(flow_row + 1):(sample_header_row - 2), 1:4]
colnames(calib) <- c("flow_setting", "weight_before", "weight_after", "time_min")
to_numeric <- function(x) {
  suppressWarnings(as.numeric(gsub(",", ".", as.character(x))))
}

calib$flow_setting <- to_numeric(calib$flow_setting)
calib$weight_before <- to_numeric(calib$weight_before)
calib$weight_after <- to_numeric(calib$weight_after)
calib$time_min <- to_numeric(calib$time_min)

calib <- calib[complete.cases(calib), ]

if (nrow(calib) < 2) {
  stop("Not enough calibration rows to fit a regression.")
}

calib$actual_flow <- (calib$weight_before - calib$weight_after) / calib$time_min * 1000
fit <- lm(actual_flow ~ flow_setting, data = calib)
coefs <- coef(fit)

samples <- read_excel(fac_path, skip = sample_header_row - 1, col_names = TRUE)
if (!"Sample Number" %in% names(samples)) {
  stop("Sample table header not found after row ", sample_header_row)
}

samples <- samples[!is.na(samples$`Sample Number`), ]
samples$`Sample Number` <- to_numeric(samples$`Sample Number`)
samples$Dilution <- to_numeric(samples$Dilution)
samples$Time <- to_numeric(samples$Time)
samples$Flow <- to_numeric(samples$Flow)
if (nrow(samples) == 0) {
  stop("No sample rows found.")
}

samples$flow_ul_min <- coefs[["flow_setting"]] * samples$Flow + coefs[["(Intercept)"]]
samples$volume_passed_ul <- samples$flow_ul_min * samples$Time / 60

flow_lookup <- data.frame(
  sample_number = samples$`Sample Number`,
  cell_line = samples$`Cell line`,
  sample = samples$Sample,
  dilution = samples$Dilution,
  time_s = samples$Time,
  flow_setting = samples$Flow,
  flow_ul_min = samples$flow_ul_min,
  volume_passed_ul = samples$volume_passed_ul,
  stringsAsFactors = FALSE
)

write.csv(flow_lookup, out_csv, row.names = FALSE)

cat("Wrote flow lookup to", out_csv, "\n")
cat("Calibration fit: y =", round(coefs[["flow_setting"]], 6), "* x +", round(coefs[["(Intercept)"]], 6), "\n")
cat("R-squared:", round(summary(fit)$r.squared, 6), "\n")
