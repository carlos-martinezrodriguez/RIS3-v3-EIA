library(readxl)
library(dplyr)
library(tidyr)
library(stringr)
library(ggplot2)

# Load and prep
load_survey <- function(path, sheet = 1){
  raw <- read_xlsx(path, sheet = sheet, col_names = FALSE)
  
  cat <- as.character(unlist(raw[1, ]))     # row 1: Category
  sub <- as.character(unlist(raw[2, ]))     # row 2: Subcategory
  
  # forward-fill blanks in row 1 (category)
  if (length(cat) > 1) {
    for (i in 2:length(cat)) {
      if (is.na(cat[i]) || cat[i] == "") cat[i] <- cat[i-1]
    }
  }
  
  header <- ifelse(is.na(sub) | sub == "", cat, paste0(cat, "_", sub))
  header[1:2] <- c("Question", "Response")  # A,B
  names(raw) <- header
  
  df <- raw[-c(1:3), ]                      # drop row1, row2, row3 (%)
  df
}

# default group set (same as your example chart)
default_groups <- function() c(
  "Total_Total",
  "Gender_Male","Gender_Female",
  "Age_17 - 24","Age_25 - 34","Age_35 - 44","Age_45 - 54",
  "Age_55 - 59","Age_60 - 69","Age_70 - 79","Age_80 - over",
  "Ethnicity_White","Ethnicity_Black, Asian, Mixed and Other",
  "Has a disability_Disability","Has a disability_No disability"
)
default_labels <- function() c(
  "Total_Total"="Total",
  "Gender_Male"="Male","Gender_Female"="Female",
  "Age_17 - 24"="17–24","Age_25 - 34"="25–34","Age_35 - 44"="35–44",
  "Age_45 - 54"="45–54","Age_55 - 59"="55–59","Age_60 - 69"="60–69",
  "Age_70 - 79"="70–79","Age_80 - over"="80+",
  "Ethnicity_White"="White",
  "Ethnicity_Black, Asian, Mixed and Other"="Ethnic Minority Groups",
  "Has a disability_Disability"="Disability",
  "Has a disability_No disability"="No disability"
)

.to_num <- function(x) as.numeric(gsub("[^0-9.\\-]", "", as.character(x)))

# keep your previous helpers: load_survey(), default_groups(), default_labels(), .to_num()

# Compare each subgroup either to its COMPLEMENT (default, statistically clean)
# or to the overall TOTAL (via a binomial test; optimistic because Total is estimated).
# You can also set a design effect and a minimum difference in percentage points.
group_sig_values <- function(df, rows, base_row,
                             groups = default_groups(),
                             labels = default_labels(),
                             total_col = "Total_Total",
                             compare_to = c("complement","total"),
                             alpha = 0.05,
                             p_adjust = "BH",
                             deff = 1,                  # e.g., 1.5 or 2 if you want to be conservative
                             min_diff_pp = 0,           # e.g., 3 to require ≥3pp difference
                             excel_rows = TRUE) {
  
  compare_to <- match.arg(compare_to)
  
  r <- if (excel_rows) rows - 3L else rows
  b <- if (excel_rows) base_row - 3L else base_row
  
  cols <- unique(c(total_col, groups))
  cols <- cols[cols %in% names(df)]
  
  perc <- df[r, cols, drop = FALSE] |>
    summarise(across(everything(), ~ sum(.to_num(.), na.rm = TRUE)))
  base <- df[b, cols, drop = FALSE] |>
    summarise(across(everything(), ~ .to_num(.)))
  
  p_total <- as.numeric(perc[[total_col]])
  n_total <- as.numeric(base[[total_col]]) / deff
  
  out <- lapply(setdiff(cols, total_col), function(g) {
    p_g <- as.numeric(perc[[g]])
    n_g <- as.numeric(base[[g]]) / deff
    
    if (is.na(p_g) || is.na(n_g) || n_g <= 0 || is.na(p_total) || is.na(n_total) || n_total <= n_g)
      return(NULL)
    
    s_g <- round(p_g/100 * n_g)
    
    if (compare_to == "complement") {
      # derive complement from totals
      p_comp_raw <- (p_total * (n_total*deff) - p_g * (n_g*deff)) / ((n_total*deff) - (n_g*deff))
      n_comp <- n_total - n_g
      s_comp <- round(p_comp_raw/100 * n_comp)
      tab <- matrix(c(s_g, round(n_g - s_g), s_comp, round(n_comp - s_comp)), nrow = 2)
      pval <- tryCatch(fisher.test(tab)$p.value, error = function(e) NA_real_)
      p_ref <- p_comp_raw
    } else {
      # subgroup vs overall total (treat total as fixed p0) -- optimistic
      pval <- tryCatch(binom.test(s_g, round(n_g), p = p_total/100)$p.value, error = function(e) NA_real_)
      p_ref <- p_total
    }
    
    data.frame(group = g,
               percent = p_g,
               n_eff = n_g,
               pval = pval,
               diff_pp = p_g - p_ref,
               direction = ifelse(p_g >= p_ref, "higher", "lower"),
               stringsAsFactors = FALSE)
  }) |> bind_rows()
  
  if (nrow(out) == 0) return(out)
  
  out$pval_adj <- if (p_adjust == "none") out$pval else p.adjust(out$pval, method = p_adjust)
  out$sig <- !is.na(out$pval_adj) & out$pval_adj < alpha & abs(out$diff_pp) >= min_diff_pp
  
  out$label <- labels[out$group]
  out$label <- factor(out$label, levels = labels[groups[groups %in% out$group]])
  out
}

# vals = output from group_sig_values(...)
# palette can be any named vector with higher/ns/lower
plot_sig_chart <- function(vals, title = "", caption = NULL,
                           palette = c(higher = "#2C7BB6",  # blue (higher)
                                       ns     = "#9ECAE1",  # light blue
                                       lower  = "#F28E2B"   # orange
                           ),
                           folder = "SRUS Graphs",
                           compare_to = c("complement","total")) {
  
  compare_to <- match.arg(compare_to)
  
  # grab object name (e.g. vals_weekly → "vals_weekly")
  obj_name <- deparse(substitute(vals))
  # strip leading "vals_" if present
  suffix <- sub("^vals_", "", obj_name)
  
  # build file name
  if (!dir.exists(folder)) dir.create(folder, recursive = TRUE)
  file_out <- file.path(folder, paste0("srus_", suffix, "_", compare_to, ".png"))
  
  vals <- vals |>
    dplyr::mutate(sig_cat = dplyr::case_when(
      sig & direction == "higher" ~ "higher",
      sig & direction == "lower"  ~ "lower",
      TRUE                        ~ "ns"
    ),
    sig_cat = factor(sig_cat, levels = c("higher","ns","lower"))
    ) |>
    dplyr::filter(!is.na(percent), percent != 0)
  
  p <- ggplot(vals, aes(x = label, y = percent, fill = sig_cat)) +
    geom_col(width = 0.7) +
    geom_text(aes(label = round(percent)), vjust = -0.35, size = 3) +
    scale_fill_manual(
      values = palette,
      breaks = c("higher","ns","lower"),
      labels = c("Higher than average", "No significant difference", "Lower than average")
    ) +
    labs(x = NULL, y = NULL, title = title, caption = caption) +
    scale_y_continuous(expand = expansion(mult = c(0, .12))) +
    theme_minimal(base_size = 11) +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          panel.grid.minor = element_blank(),
          legend.title = element_blank())
  
  ggsave(file_out, plot = p, width = 8, height = 5, dpi = 300)
  message("Saved: ", file_out)
  
  invisible(p)
}

# Row collapse and plotting

#Frequency of use
vals_weekly <- group_sig_values(df, rows = 4:6, base_row = 11,
                                compare_to = "total",
                                deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_weekly,
               title = "% Using the SRN at Least Once Weekly (Significance against Total)", compare_to = "total")

vals_weekly <- group_sig_values(df, rows = 4:6, base_row = 11,
                                compare_to = "complement",
                                deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_weekly,
               title = "% Using the SRN at Least Once Weekly (Significance against Complement)", compare_to = "complement")

#Annual mileage 
vals_mileage <- group_sig_values(df, rows = 15:18, base_row = 21,
                                 compare_to = "total",
                                 deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_mileage,
               title = "% With More than 10k Annual Miles (Significance against Total)", compare_to = "total")

vals_mileage <- group_sig_values(df, rows = 15:18, base_row = 21,
                                 compare_to = "complement",
                                 deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_mileage,
               title = "% With More than 10k Annual Miles (Significance against Complement)", compare_to = "complement")

#Purpose: Commuting
vals_purpose_commute <- group_sig_values(df, rows = 23, base_row = 31,
                                         compare_to = "total",
                                         deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_purpose_commute,
               title = "% Use SRN for Commuting (Significance against Total)", compare_to = "total")

vals_purpose_commute <- group_sig_values(df, rows = 23, base_row = 31,
                                         compare_to = "complement",
                                         deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_purpose_commute,
               title = "% Use SRN for Commuting (Significance against Complement)", compare_to = "complement")

#Purpose: Leisure
vals_purpose_leisure <- group_sig_values(df, rows = 24, base_row = 31,
                                         compare_to = "total",
                                         deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_purpose_leisure,
               title = "% Use SRN for Leisure (Significance against Total)", compare_to = "total")

vals_purpose_leisure <- group_sig_values(df, rows = 24, base_row = 31,
                                         compare_to = "complement",
                                         deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_purpose_leisure,
               title = "% Use SRN for Leisure (Significance against Complement)", compare_to = "complement")


# Overall satisfaction
vals_satisfaction <- group_sig_values(df, rows = 32:33, base_row = 39,
                                         compare_to = "total",
                                         deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_satisfaction,
               title = "% Very or Fairly Satisfied with the SRN (Significance against Total)", compare_to = "total")

vals_satisfaction <- group_sig_values(df, rows = 32:33, base_row = 39,
                                      compare_to = "complement",
                                      deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_satisfaction,
               title = "% Very or Fairly Satisfied with the SRN (Significance against Complement)", compare_to = "complement")


#Journey Time Satisfaction
vals_jts <- group_sig_values(df, rows = 48:49, base_row = 55,
                                      compare_to = "total",
                                      deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_jts,
               title = "% Very or Fairly Satisfied with Journey Time (Significance against Total)", compare_to = "total")

vals_jts <- group_sig_values(df, rows = 48:49, base_row = 55,
                             compare_to = "complement",
                             deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_jts,
               title = "% Very or Fairly Satisfied with Journey Time (Significance against Complement)", compare_to = "complement")


# Satisfaction with Roadworks
vals_satisfaction_roadworks <- group_sig_values(df, rows = 56:59, base_row = 70,
                             compare_to = "total",
                             deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_satisfaction_roadworks,
               title = "% Very or Fairly Satisfied with Roadwork Management (Significance against Total)", compare_to = "total")

vals_satisfaction_roadworks <- group_sig_values(df, rows = 56:59, base_row = 70,
                                                compare_to = "complement",
                                                deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_satisfaction_roadworks,
               title = "% Very or Fairly Satisfied with Roadwork Management (Significance against Complement)", compare_to = "complement")


# Feelings of Safety
vals_feel_safe <- group_sig_values(df, rows = 40:41, base_row = 47,
                                                compare_to = "total",
                                                deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_feel_safe,
               title = "% Feel Safe or Very Safe when Using the SRN (Significance against Total)", compare_to = "total")

vals_feel_safe <- group_sig_values(df, rows = 40:41, base_row = 47,
                                   compare_to = "complement",
                                   deff = 1.7, min_diff_pp = 3)
plot_sig_chart(vals_feel_safe,
               title = "% Feel Safe or Very Safe when Using the SRN (Significance against Complement)", compare_to = "complement")


