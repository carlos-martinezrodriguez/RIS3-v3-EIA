library(readxl)
library(dplyr)
library(ggplot2)
library(stringr)

plot_disability_car_access_pct <- function(
    path = "NTS/Surveys/NTS trips by personal car access, disability status and age.xlsx",
    sheet = "DIS0405b",
    year_contains = "2024",                     # matches "2023 to 2024", etc.
    outdir = "NTS/NTS Graphs",
    outfile = "nts_disability_car_access_2024.png",
    min_diff_pp = 3                             # require >= 3pp for highlighting
){
  df <- read_xlsx(path, sheet = sheet, skip = 4)
  
  # Keep the year window that contains 2024 and the All-ages rows
  df_yr <- df %>%
    filter(str_detect(Year, fixed(year_contains)),
           `Age group` == "All ages 16 and over") %>%
    select(
      `Disability status`,
      percent = `Adults in households with a car or van (%)`,
      n_raw   = `Unweighted sample size: individuals [notes 4, 6] (number)`
    ) %>%
    mutate(
      percent = as.numeric(percent),
      n       = as.numeric(gsub(",", "", n_raw))
    )
  
  # overall total as weighted mean across Disabled + Not disabled
  p_total <- with(df_yr, sum(percent * n, na.rm = TRUE) / sum(n, na.rm = TRUE))
  
  # Significance vs total (binomial, using unweighted bases from the sheet)
  results <- df_yr %>%
    filter(!is.na(percent), !is.na(n), n > 0) %>%
    rowwise() %>%
    mutate(
      s_g     = round(percent/100 * n),
      pval    = tryCatch(binom.test(s_g, n, p = p_total/100)$p.value, error = function(e) NA_real_),
      diff_pp = percent - p_total,
      direction = ifelse(percent > p_total, "higher", "lower"),
      sig = !is.na(pval) & pval < 0.05 & abs(diff_pp) >= min_diff_pp
    ) %>% ungroup()
  
  # Plot (three-state colouring: higher / ns / lower)
  p <- ggplot(results, aes(x = `Disability status`, y = percent,
                           fill = case_when(
                             sig & direction == "higher" ~ "higher",
                             sig & direction == "lower"  ~ "lower",
                             TRUE                        ~ "ns"
                           ))) +
    geom_col(width = 0.6) +
    geom_text(aes(label = round(percent)), vjust = -0.3, size = 3) +
    scale_fill_manual(
      breaks = c("higher","ns","lower"),
      values = c(higher="#2C7BB6", ns="#9ECAE1", lower="#F28E2B"),
      labels = c("Higher than total", "No significant difference", "Lower than total"),
      name = NULL
    ) +
    labs(x = NULL, y = "% in households with a car/van",
         title = "Car/van access by disability status (All ages 16+; 2024)",
         caption = "Compared to overall (weighted by base). Binomial test p<0.05; ≥3pp difference.") +
    scale_y_continuous(limits = c(0, 100)) +
    theme_minimal(base_size = 11) +
    theme(panel.grid.minor = element_blank(),
          axis.text.x = element_text(angle = 0, hjust = 0.5))
  
  if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)
  outfile_path <- file.path(outdir, outfile)
  ggsave(outfile_path, plot = p, width = 6, height = 4, dpi = 300)
  message("Saved to: ", normalizePath(outfile_path, winslash = "/"))
  
  p
}

# Run it
plot_disability_car_access_pct()

